This folder contains a UFT project used as a driver for Functional Testing.
It is designed to be reusable and produce repeatable results.
Most of the Actions are primarily shared Actions that are re-used.
The Actions are data-driven by XLS datasheets.
The goal of each Action is to import SPT numbers and check for an expected result.
The Default.XLS is mostly empty and must be loaded at runtime via other XLS files.

The ProvisionNew.xls verifies attempts to import new SPT numbers.
The ProvisionUsed.xls verifies attempts to import used (consumed) SPT numbers.
The PreLot_*.xls import new SPT numbers (preload data before running a lot).
The PostLot_*.xls disables any remaining SPT numbers (prep for next lot; cleanup of unused numbers).

