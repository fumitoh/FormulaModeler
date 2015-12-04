++++++++++++++++++++
Formula Modeler News
++++++++++++++++++++

What's New in Formula Modeler 0.4.0?
====================================

*Release date: 21-Sep-2014*

Core
----

- Added VecToLine method to FModel.

Library
-------

- Projection2, the yearly projction formula is added.
- Added gprj_InvRetRate and gprj_DiscRate in GlobalProjection formula.
- Yearly Deterministic Projection Model is added.
- Changed Runcontrol2 to output by-policy and total results using VecToLine.
- In Projection formula, variable names ending _EoM or _BoM changed to _EoP or _BoP. 

Documentation
-------------

- Updated due to the introduction of Projection2, Yearly Deterministic Projection Model.
- Updated due to the introeuction of FModel::VecToLine.

What's New in Formula Modeler 0.3.1?
====================================

*Release date: 14-Sep-2014*

Core
----

- Bug Fix: Fixed errors that happen when VBA compile command is executed.
- Bug Fix: MultDimArray now accepts a valid 1 dimentional array as Source.

Library
-------

- actrl_ExcelHPC module is now included in the distribution file as 
  a .bas file separately from the library file, fml_actuarial_model_0_3_1.xlsm.
- Got rid of a Japanese error message in actrl_Run::Run_Model.

Documentation
-------------

- Syntax highlighting is enabled on the VBA code in the manuals.
- Revised MultDimArray section in User Reference to describe the function more accurately.
- Changed the default font to Meiryo.
- Revised the description on running models with HPC.
- Added the explanation of sequence diagrams in the Introduction section.

