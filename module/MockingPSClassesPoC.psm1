using namespace System.Collections.Generic
using namespace System.Management.Automation
using namespace System.Management.Automation.Internal
using namespace System.Management.Automation.Language
using namespace System.Reflection

# Represents the conditions and filters of a single mocked method.  For our purposes a single method
# is defined by type name and method name.  All overloads look to the same mock table, but can be
# selected individually with parameter filters.
class MethodMock {
    [List[Tuple[scriptblock, scriptblock]]] $Filters;
    [string] $Name;

    MethodMock([string] $name) {
        $this.Filters = [List[Tuple[scriptblock, scriptblock]]]::new()
        $this.Name = $name
    }

    [void] AddCondition([scriptblock] $condition, [scriptblock] $mock) {
        $this.Filters.Insert(
            0,
            [Tuple[scriptblock, scriptblock]]::new(
                $condition,
                $mock))
    }
}

# Provides an central store for mock controls and conditions
class MockFactory {
    # Contains type names that are targetted for mocks
    static [HashSet[string]] $MockedTypes;

    # Contains types that have been initialized
    static [HashSet[type]] $InitializedTypes;

    # Contains type names that are targetted for mocks but have not yet been loaded
    static [HashSet[string]] $InitializeOnLoad;

    # Contains a field to original member wrapper mapping for every method initialized for mocking
    static [Dictionary[FieldInfo, ScriptBlockMemberMethodWrapper]] $ResetTable;

    # Contains mock conditions and actions for a method group
    static [Dictionary[string, MethodMock]] $MockTable;

    static MockFactory() {
        $ignoreCase = [StringComparer]::CurrentCultureIgnoreCase
        [MockFactory]::MockedTypes = [HashSet[string]]::new($ignoreCase)
        [MockFactory]::InitializedTypes = [HashSet[type]]::new()
        [MockFactory]::InitializeOnLoad = [HashSet[string]]::new($ignoreCase)
        [MockFactory]::ResetTable = [Dictionary[FieldInfo, ScriptBlockMemberMethodWrapper]]::new()
        [MockFactory]::MockTable = [Dictionary[string, MethodMock]]::new($ignoreCase)
    }

    # Replaces all script block wrappers with ones we control. Wrapped methods aren't actually
    # mocked until a condition is added to the mock table.
    static [void] InitializeTypeForMocking([string] $typeName) {
        $types = [ReflectionOps]::GetType($typeName)
        if (-not $types.Count) {
            # If we can't find any loaded types then add the type to the watch list
            # for the assembly load event
            [MockFactory]::InitializeOnLoad.Add($typeName)
            return
        }

        if ([MockFactory]::InitializeOnLoad.Contains($typeName)) {
            [MockFactory]::InitializeOnLoad.Remove($typeName)
        }

        foreach ($type in $types) {
            if ([MockFactory]::InitializedTypes.Contains($type)) {
                continue
            }

            [MockFactory]::MockTypeMembers($type)
        }

        [MockFactory]::MockedTypes.Add($typeName)
    }

    # Called by the assembly load event subscriber to mock classes as they load
    static [void] MaybeMockAssembly([Assembly] $assembly) {
        if ($assembly.CustomAttributes.AttributeType -notcontains [ReflectionOps]::DynamicAssemblyMarker) {
            return
        }

        $assembly.GetTypes().ForEach{ [MockFactory]::MockTypeMembers($PSItem) }
    }

    # Called by the assembly load event subscriber to mock classes that were either
    # not available during the initial mock or if a new version is loaded
    static [void] MaybeMockTypeMembers([type] $type) {
        if ($type.Name.Contains('<staticHelpers>')) {
            return
        }

        if ([MockFactory]::InitializeOnLoad.Contains($type.Name)) {
            [MockFactory]::InitializeTypeForMocking($type.Name)
            return
        }

        if (-not [MockFactory]::MockedTypes.Contains($type.Name)) {
            return
        }

        [MockFactory]::MockTypeMembers($type)
    }

    
    static [void] MockTypeMembers([type] $type) {
        if ($type.Name.Contains('<staticHelpers>')) {
            return
        }

        # Every PSClass with methods has a non-public type called 'TypeName_<staticHelpers>' that
        # houses the script blocks invoked by class methods.
        $staticHelper = [ReflectionOps]::GetStaticHelperForType($type)

        # If we couldn't find a static helper class it's probably an enum
        if (-not $staticHelper) {
            return
        }
        
        # The static helper has a field for each method that contains a script block wrapper
        foreach ($memberField in [ReflectionOps]::GetMethodFields($staticHelper)) {
            # If it's in the reset table we've already taken control of this method.
            if ([MockFactory]::ResetTable.ContainsKey($memberField)) {
                continue
            }

            $wrapper = $memberField.GetValue($null)
            # Skip compiler generated constructors from default property values
            if ([ReflectionCache]::Wrapper_Ast.GetValue($wrapper) -isnot [FunctionMemberAst]) {
                continue
            }

            # A mock contains the parameter filters and "mock with"'s. All methods in a type
            # are mocked at the same time even if there is only a mock for one method.  If no
            # mock conditions exist in the mock then the original method will be invoked.
            [MockFactory]::GetOrCreateMock($wrapper)
            $newWrapper = [MethodUtil]::CreateMockedWrapper(
                $wrapper,
                [ReflectionOps]::GetMetadataPath($memberField))

            $memberField.SetValue($null, $newWrapper)

            # Add the field and original wrapper to the reset table so we can:
            #  1. Fall back to the original method if no conditions exist
            #  2. Revert to the original wrapper if TearDown is called
            [MockFactory]::ResetTable.Add($memberField, $wrapper)
        }

        [MockFactory]::InitializedTypes.Add($type)
    }

    # Retrieve a mock from the mock table or create a new one if not available
    static [MethodMock] GetOrCreateMock([ScriptBlockMemberMethodWrapper] $wrapper) {
        return [MockFactory]::GetOrCreateMock(
            [MethodUtil]::GetNameFromWrapper($wrapper))
    }

    # Retrieve a mock from the mock table or create a new one if not available
    static [MethodMock] GetOrCreateMock([string] $name) {
        $mock = $null
        if ([MockFactory]::MockTable.TryGetValue($name, [ref]$mock)) {
            return $mock
        }

        $mock = [MethodMock]::new($name)
        [MockFactory]::MockTable.Add($name, $mock)

        return $mock
    }

    # Resolve a metadata path to retrieve the associated mock, then evaluate parameter filters to
    # determine the script block to return. This method is called from the mocked method to determine
    # what to invoke.
    static [scriptblock] EvaluateAndGetMock([string] $metadataPath, [SessionState] $state) {
        # A string is injected into the mocked method containing a resolvable path to the field
        # for the method invoked. This is done because as far as I can tell there is no way to
        # determine what method you are currently in.  InvocationInfo doesn't even have a CommandInfo
        # object to evaluate.  Even the FunctionContext doesn't have a script block, so instead a
        # path of metadata tokens injected into the AST while creating the mock wrapper.
        $field = [ReflectionOps]::ResolveMetadataPath($metadataPath)
        $mock = [MockFactory]::GetOrCreateMock($field.GetValue($null))
        foreach ($condition in $mock.Filters) {
            if (. $condition.Item1) {
                return $condition.Item2
            }
        }

        # None of the conditions passed (or there weren't any) so fall back to the original method.
        $originalWrapper = [MockFactory]::ResetTable[$field]

        # Recreating the initial script block isn't ideal, but parameter variables aren't visible
        # without it.  Consider changing this to use the InvokeHelper method on the wrapper.
        $originalSb = [ReflectionCache]::Wrapper_BoundScriptBlock.
            GetValue($originalWrapper).Value.Ast.Body.
            GetScriptBlock()

        $internal = [ReflectionCache]::SessionState_Internal.GetValue($state)
        [ReflectionCache]::ScriptBlock_SessionStateInternal.SetValue(
            $originalSb,
            $internal)

        return $originalSb
    }

    # Reset all mocks and tables to their initial state.
    static [void] TearDown() {
        foreach ($pair in [MockFactory]::ResetTable.GetEnumerator()) {
            $pair.Key.SetValue($null, $pair.Value)
        }

        [MockFactory]::ResetTable.Clear()
        [MockFactory]::MockTable.Clear()
        [MockFactory]::MockedTypes.Clear()
        [MockFactory]::InitializeOnLoad.Clear()
        [MockFactory]::InitializedTypes.Clear()
    }
}

# Provides utility methods for working with script block wrappers
class MethodUtil {
    static [scriptblock] $BodyReturnsObject;
    static [scriptblock] $BodyReturnsVoid;

    static MethodUtil() {
        # Need a explicit return statement if the method returns
        [MethodUtil]::BodyReturnsObject = {
            return . ([MockFactory]::EvaluateAndGetMock('{0}', $ExecutionContext.SessionState))
        }

        [MethodUtil]::BodyReturnsVoid = {
            . ([MockFactory]::EvaluateAndGetMock('{0}', $ExecutionContext.SessionState))
        }
    }

    static [string] GetNameFromWrapper([ScriptBlockMemberMethodWrapper] $wrapper) {
        $ast = [MethodUtil]::GetAstFromWrapper($wrapper)
        return $ast.Parent.Name + '\' + $ast.Name
    }

    # Create a new script block wrapper that we can add mock conditions to
    static [ScriptBlockMemberMethodWrapper] CreateMockedWrapper(
        [ScriptBlockMemberMethodWrapper] $wrapper,
        [string] $metadataPath)
    {
        $ast = [MethodUtil]::GetAstFromWrapper($wrapper)
        $body = [MethodUtil]::GetMockBody($ast, $metadataPath)

        return [ReflectionCache]::Wrapper_Ctor.Invoke(@(
            [MethodUtil]::CloneMemberAstWithNewBody($ast, $body)))
    }

    # The body must explicitly return if the method has a return value, otherwise it must _not_
    # explicitly return.  We also inject a metadata path so the method can resolve itself.
    static [ScriptBlockAst] GetMockBody([MemberAst] $ast, [string] $metadataPath) {
        return [Parser]::ParseInput(
            [MethodUtil]::GetMockBodyText($ast) -f $metadataPath,
            [ref]$null,
            [ref]$null)
    }

    static [string] GetMockBodyText([MemberAst] $ast) {
        if ($null -eq $ast.ReturnType -or $ast.ReturnType -eq [void]) {
            return [MethodUtil]::BodyReturnsVoid
        }

        return [MethodUtil]::BodyReturnsObject
    }

    static [MemberAst] GetAstFromWrapper([ScriptBlockMemberMethodWrapper] $wrapper) {
        return [ReflectionCache]::Wrapper_Ast.GetValue($wrapper)
    }

    # Recreate a MemberAst with a new body.  This way it keeps extent data, parameters, return type,
    # name, and the rest of the TypeDefinitionAst.
    static [FunctionMemberAst] CloneMemberAstWithNewBody([FunctionMemberAst] $source, [ScriptBlockAst] $body) {
        $statements = [StatementBlockAst]::new(
            <# extent:     #> $source.Body.EndBlock.Extent,
            <# statements: #> $body.EndBlock.Statements.ForEach('Copy').ForEach([StatementAst]),
            <# traps:      #> [TrapStatementAst[]]::new(0))
        
        $sbAst = [ScriptBlockAst]::new(
            <# extent:     #> $source.Body.Extent,
            <# paramBlock: #> $source.Body.ParamBlock.ForEach('Copy')[0],
            <# statements: #> $statements,
            <# isFilter:   #> $false)

        $functionAst = [FunctionDefinitionAst]::new(
            <# extent:     #> $source.Extent,
            <# isFilter:   #> $false,
            <# isWorkflow: #> $false,
            <# name:       #> $source.Name,
            <# parameters: #> $source.Parameters.ForEach('Copy').ForEach([ParameterAst]),
            <# body:       #> $sbAst)

        $memberAst = [FunctionMemberAst]::new(
            <# extent:                #> $source.Extent,
            <# functionDefinitionAst: #> $functionAst,
            <# returnType:            #> $source.ReturnType.ForEach('Copy')[0],
            <# attributes:            #> $source.Attributes.ForEach('Copy').ForEach([AttributeAst]),
            <# methodAttributes:      #> $source.MethodAttributes)

        [ReflectionCache]::Ast_SetParent.Invoke($source.Parent, @($memberAst))

        return $memberAst
    }
}

# Provides utility methods for finding PSClasses and their internals in the AppDomain
class ReflectionOps {
    static [type] $DynamicAssemblyMarker = [DynamicClassImplementationAssemblyAttribute]
    static [string] $StaticHelperName = '{0}_<staticHelpers>'

    # Get the non-public static helper for the type that houses the script block wrappers
    static [type] GetStaticHelperForType([type] $type) {
        return $type.Assembly.GetType([ReflectionOps]::StaticHelperName -f $type.ToString())
    }

    # Get all versions of a PSClass type name, including those not currently resolvable
    static [type[]] GetType([string] $typeName) {
        # Assembly.GetType would be faster, but duplicate type names in the same assembly are possible
        # if the same class is defined in multiple scopes in the same file.
        return [AppDomain]::CurrentDomain.GetAssemblies().
            Where{ $PSItem.CustomAttributes.AttributeType -contains [ReflectionOps]::DynamicAssemblyMarker }.
            ForEach{ $PSItem.GetTypes() }.
            Where{ $PSItem.Name -eq $typeName }.
            ForEach([type])
    }

    # Get the fields the house the script block wrappers from a static helper type
    static [FieldInfo[]] GetMethodFields([type] $staticHelper) {
        return $staticHelper.
            GetFields([BindingFlags]'Static, NonPublic').
            Where{ $PSItem.FieldType -eq [ScriptBlockMemberMethodWrapper] }.
            ForEach([FieldInfo])
    }

    # Build a short "path" of metadata tokens and hash codes that we can use to resolve
    # a specific field from a string.
    static [string] GetMetadataPath([FieldInfo] $field) {
        return '' +
            $field.ReflectedType.Assembly.GetHashCode() + '.' +
            $field.ReflectedType.Module.MetadataToken + '.' +
            $field.MetadataToken
    }

    # Resolve a metadata path
    static [FieldInfo] ResolveMetadataPath([string] $path) {
        $assemblyId, $moduleId, $fieldId = $path.Split('.')
        return [AppDomain]::
            CurrentDomain.
            GetAssemblies().
            Where({ $PSItem.GetHashCode() -eq $assemblyId }, 'First')[0].
            GetModules().
            Where({ $PSItem.MetadataToken -eq $moduleId }, 'First')[0].
            ResolveField($fieldId)
    }
}

# Provides a central store of non-public members for performance and some improvement to readability
class ReflectionCache {
    static [BindingFlags] $InstanceFlags = [BindingFlags]'NonPublic, Instance';
    static [BindingFlags] $StaticFlags = [BindingFlags]'NonPublic, Static';
    static [type] $IParameterMetadataProvider = [ref].Assembly.GetType('System.Management.Automation.Language.IParameterMetadataProvider');
    static [type] $TypeAccelerators = [ref].Assembly.GetType('System.Management.Automation.TypeAccelerators');
    static [FieldInfo] $Wrapper_Ast = [ScriptBlockMemberMethodWrapper].GetField('_ast', [ReflectionCache]::InstanceFlags);
    static [FieldInfo] $Wrapper_BoundScriptBlock = [ScriptBlockMemberMethodWrapper].GetField('_boundScriptBlock', [ReflectionCache]::InstanceFlags);
    static [PropertyInfo] $SessionState_Internal = [SessionState].GetProperty('Internal', [ReflectionCache]::InstanceFlags);
    static [PropertyInfo] $ScriptBlock_SessionStateInternal = [scriptblock].GetProperty('SessionStateInternal', [ReflectionCache]::InstanceFlags);
    static [MethodInfo] $Ast_SetParent = [Ast].GetMethod('SetParent', [ReflectionCache]::InstanceFlags);
    static [ConstructorInfo] $Wrapper_Ctor = [ScriptBlockMemberMethodWrapper].GetConstructor([ReflectionCache]::InstanceFlags, $null, @([ReflectionCache]::IParameterMetadataProvider), 1);
}

function Add-MethodMock {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $TypeName,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $MethodName,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [scriptblock] $MockWith,

        [ValidateNotNull()]
        [scriptblock] $ParameterFilter = { $true }
    )
    end {
        [MockFactory]::InitializeTypeForMocking($TypeName)
        $mock = [MockFactory]::GetOrCreateMock("$TypeName\$MethodName")
        $mock.AddCondition($ParameterFilter, $MockWith)
    }
}

function Clear-MethodMock {
    [CmdletBinding()]
    param()
    end {
        [MockFactory]::TearDown()
    }
}

# A type accelerator is added so MockFactory can be resolved in any scope/runspace
[ReflectionCache]::TypeAccelerators::Add(
    'MockFactory',
    [MockFactory])

# Subscribe to AssemblyLoad events so we can mock new revisions of mocked classes added
# at runtime. This also allows us to add a mock prior to initially loading the module that
# contains the class as well.
$SUBSCRIBER_GUID = [guid]::NewGuid().ToString()
& {
    $registerObjectEventSplat = @{
        SourceIdentifier = $SUBSCRIBER_GUID
        InputObject      = [AppDomain]::CurrentDomain
        EventName        = 'AssemblyLoad'
        Action           = { [MockFactory]::MaybeMockAssembly($eventArgs.LoadedAssembly) }
        SupportEvent     = $true
    }

    $null = Register-ObjectEvent @registerObjectEventSplat
}

# Clear all mocks and remove the subscriber when the module is removed
$ExecutionContext.SessionState.Module.OnRemove = {
    [MockFactory]::TearDown()
    $ExecutionContext.Events.GetEventSubscribers($SUBSCRIBER_GUID).ForEach{
        $ExecutionContext.Events.UnsubscribeEvent($PSItem)
    }
}

Export-ModuleMember -Function Add-MethodMock, Clear-MethodMock
