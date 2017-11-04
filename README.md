# PSClass Mock Proof of Concept

This module is a proof of concept for a potential way to add PowerShell class mocking functionality for
test scripts.  The demo is fully functional, though I have not done any extensive tests, only enough
to know it works.  This module is intended to be a demonstration only.

If anyone wants to build this into a proper implementation feel free to reach out if you run into issues.

## Pros

- Fully mock any method in any PowerShell class, both static and instance
- Mocks apply to existing instances as well as new instances
- Mocks apply to all classes of the same name, even stale versions in older scopes
- Mocks can be defined for classes that haven't been defined yet
- Supports parameter filters
- Mocks will persist even if called by compiled classes

## Cons

- The way mocking is done uses a significant amount of reflection to access PowerShell internals
- Properties cannot be mocked.  This can be worked around by mocking the constructor
- Compiler generated functions cannot be mocked.  These are created when you give a property a default
  value.  This can be worked around by defining a constructor

## How

When a PowerShell class is defined there is another class created with it called a "static helper".
This class contains a `ScriptBlockMemberMethodWrapper` object for each method defined in the class.
Mocking is possible by replacing that wrapper with a new one that determines the method it was called
from, evaluates parameter filters, and invokes the appropriate script block.

When a mock is defined, the AppDomain is searched for matching PowerShell classes and replaces all
member wrappers with wrappers controlled by the module.  An `AssemblyLoad` event subscriber is registered
to handle the mocking of classes loaded after the fact.  Because all of this is handled via static fields
it doesn't matter when the class was loaded, what scope it was in, or what instances were already
created.

## Trying it out

```powershell
git clone https://github.com/SeeminglyScience/MockingPSClassesPoC
Import-Module ./MockingPSClassesPoC/module/MockingPSClassesPoC.psd1

Add-MethodMock MyClass MyMethod { return $this.MyProperty -replace 'not ' }
class MyClass {
    [string] $MyProperty;

    MyClass() {
        $this.MyProperty = 'NotMocked'
    }

    [string] MyMethod() {
        return $this.MyProperty
    }
}

$instance = [MyClass]::new()
$instance.MyMethod()
# Mocked
Clear-MethodMock
$instance.MyMethod()
# Not Mocked

. {
    class MyClass {
        [string] $MyProperty;

        MyClass() {
            $this.MyProperty = 'Definitely Not Mocked'
        }

        [string] MyMethod([string] $parameter) {
            return $this.MyProperty
        }
    }

    $newInstance = [MyClass]::new()
}

Add-MethodMock MyClass MyMethod { return $this.MyProperty -replace 'not ' }
Add-MethodMock MyClass MyMethod { return 'Completely Mocked' } { $parameter -eq 'super mocked' }

$newInstance.MyMethod('default mock')
# Definitely Mocked
$newInstance.MyMethod('super mocked')
# Completely Mocked

# Mock applies to the old instance too
$instance.MyMethod()
# Mocked

Add-MethodMock MyClass MyClass { $this.MyProperty = 'Mocked from the start' }
[MyClass]::new()
# MyProperty
# ----------
# Mocked from the start
```
