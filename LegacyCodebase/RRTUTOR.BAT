@echo off
cls
echo ษอออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออป
echo บ          R&R for Xbase Tutorial Program                                     บ
echo บ          ฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤ            บ
echo บ                                                                             บ
echo บ          This batch program is intended for use in doing the R&R            บ
echo บ          tutorial exercises in the "Learning by Example" manual.            บ
echo บ          It executes R&R using a customized configuration file              บ
echo บ          (RRTUTOR.CNF) set up specifically for the tutorial                 บ
echo บ          exercises.                                                         บ
echo บ                                                                             บ
echo บ          Also note that RRTUTOR.FIF and RRTUTOR.TXT, the font               บ
echo บ          information file and ASCII text file installed with                บ
echo บ          the other tutorial files, are intended for use only                บ
echo บ          with the tutorials.                                                บ
echo บ                                                                             บ
echo บ          If you do not want to continue, press Ctrl-Break now.              บ
echo บ          Otherwise, press any key to run R&R.                               บ
echo บ                                                                             บ
echo บ                                                                             บ
echo ศอออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผ
pause 
rr.exe .\rrtutor/c
