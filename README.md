# VBScript IsEmpty Function Bug

This repository demonstrates a common issue encountered when using the `IsEmpty` function in VBScript.  The `IsEmpty` function doesn't reliably distinguish between an uninitialized variant and an empty string.  This can lead to unexpected errors and inconsistent program behavior.

## Bug Description:

The provided `bug.vbs` script showcases a situation where `IsEmpty` fails to correctly identify an empty string parameter, causing an error.  The solution in `bugSolution.vbs` offers a more robust approach.

## How to reproduce the bug:

1. Save the code in `bug.vbs`. 
2. Run the script from the command line or using a VBScript interpreter.

## Solution:

The `bugSolution.vbs` file provides a corrected version.  Instead of relying solely on `IsEmpty`, it explicitly checks for both `IsEmpty` and the length of the string.