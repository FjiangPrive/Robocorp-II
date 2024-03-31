# robocorp-certification-level-ii
RoboCorp Certification Level II

Some notes and issues that came up during the course:

- ``devdata/env`` and the reference to the ``vault.json`` file: It seems that only absolute paths do seem to work here. Neither Robot Framework ``${CURDIR}`` nor Unix approaches such as ``$HOME``, ``~``,``../vault.json`` et al work. This means that the current version of this file __will__ __not__ __work__ out of the box in a different environment without some minor changes. Additionally, if you open the file in VSCode, my venv settings will kick in (see file ``.vscode/settings.json`` and config line ``"python.pythonPath": "/Users/jsl/.robocorp/live/7b3eba72202108b9/bin/python3"``)
- I have to use VSCode as my employer does not permit the installation of the Robocorp IDE. So I generated the RPA skeleton code in VSCode. The generator does create all files but assigns an older / deprecated ``rpaframework`` version to the ``conda.yaml`` file (9.xxx). With this defaullt setting, the  PDF ``append`` option does not work as the keyword will not recognise this parameter. See comments in the source code. --> ``- rpaframework==10.6.0 # https://rpaframework.org/releasenotes.html`` in the ``conda.yaml`` file will do the trick. You did add this to the instructions (``RPA Framework keywords are not recognized by the IDE!``) but you should also mention that parameters might not get recognised if you don't bump up the version.
- With the current RPA.PDF library, I do not see any way to EMBED the image in the content of the __first__ page (where the order data is). Have a look at the python code that the keyword is based on - whatever you supply as external files (images or pdfs) will __always__ be added as a SEPARATE page in the pdf file. See also my comments in the ``tasks.robot`` file (with reference to ``https://github.com/robocorp/rpaframework/blob/master/packages/pdf/src/RPA/PDF/keywords/document.py``)

With the hardcoded paths to the vault (and the venv settings), the exam constraint ``Verify that it is possible to run your robot without manual setup`` might not be met. However, as I cannot bypass the absolute environment requirements even for the vault file, the test will not work out of the box - unless you change these settings.

All other file names, paths etc have been stored as parameters in the file. Sorry mates - I mainly use Robot Framework for backend API testing, so some approaches in the test might be a little bit clumsy. But the test itself works fine.

Improvement suggestions:

I'd like to see infos on how the program can enforce the user to enter some data when using rpa.dialogs and input boxes. WUKS loops cannot be the solution here? I don't detect this potential issue in the program so this is definitely an "out of scope".

Thanks for the challenge - it was fun!