# VBScript Range Check Function Bugs

This repository demonstrates two common but subtle bugs in a simple VBScript function designed to check if a number falls within a specified range.  The first bug stems from implicit type conversion of string inputs, and the second from the lack of null value handling.

## Bugs:

1. **Implicit Type Conversion:** The original function doesn't explicitly check the input type, leading to unexpected results when strings are passed as arguments that VBScript can implicitly convert to numbers.
2. **Null Value Handling:** The function doesn't handle null or empty values correctly, resulting in type mismatch errors.

## Solution:

The solution implements explicit type checking and error handling to address these issues. The improved function explicitly checks for numeric inputs and handles null values gracefully, ensuring robustness and predictability.