This repository demonstrates a subtle bug in VBScript related to type mismatch errors that can occur when performing arithmetic operations involving empty strings. The bug arises from how VBScript implicitly handles empty strings in such contexts, which differs from explicit null checks.

The `bug.vbs` file contains the problematic code.  The `bugSolution.vbs` file provides a corrected version that addresses the issue using explicit type checking and error handling.