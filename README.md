# Unexpected IsEmpty() Behavior in VBScript

This repository demonstrates a potential issue with VBScript's `IsEmpty()` function.  `IsEmpty()` returns `True` for both empty strings and `Null` values. This can cause problems if you're not handling these cases differently and expect different actions for each scenario.  The provided example shows how this can happen and a safe solution for handling empty parameter values.  The solution emphasizes checking for both `Null` and empty string values to maintain expected behavior. 

## How to reproduce

1. Run the `bug.vbs` script.
2. Observe how the function handles empty strings versus null values.