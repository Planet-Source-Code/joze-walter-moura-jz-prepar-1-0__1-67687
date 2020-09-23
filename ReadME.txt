RECENTS FILES & PREFERENCES MANAGER
Programmers Comments
-----------------------------------
GENERAL.
1.




This Module modPrePar.Bas can be added your Project to offer
3 features:

1. A INI File i-o dispositive;
2. A Recent Files manager;
3. A your App Parameters and Preferences at Menu manager;

In Bas Module, code is organized mode you can discard some
lines if you dont need all 3 features, e.g., you can
strip Recent Files code if you don't use it.

There is 2 template Forms:  frmRecents and frmPrePar.

The objective of separated forms was the same: optimize
discards but you will see no complex code.

If you use all features, you can merge estrategical
code in same place your form.
For example:
  the frmRecents has some code to put in Form_Load Sub;
  the frmPrePar has some code, too, same place.
  In your Form, in Form_load, you may copy one past other,
  and eliminate a few 3 redundant lines (abou INIFile path).

PREMISSES:
---------
a) The Module Bas contains all the code. Forms are very small
references to do, mode don't pollute your form with solved code.

b) 20% of code is not mine of all: I have collected it by PSC
authors, so I would like to give them equal credits and rights.
Sorry, I've no way to name each of them.

c) There are many routines, variables, etc, prefixed "Jz".
In fact, it becames from Joze, my name, but the ideal was
avoid Your Project conflicts, e.g., WriteINI is a very common
name - so I use JzWriteINI to make difference.
The letters "Jz" can be globally changed by any if you want.

NEW FEATURE IN INI FILE:
-----------------------
INI file is the same of all, but lines may have Remarks, after
last ";" caracter, if present. Ex:

[Persons]
Commander=John Smith ; Who is the Boss?

So, the INI routines will work with Public variables, as:

JzINIFile = complete path/filename for INI file.
JzSection = Section name, e.g., "Persons"
JzEntry = Key name, e.g., "Commander"
JzValuINI = complete entry string, e.g.,
            "John Smith ; Who is the Boss?"
JzValue = True entry, e.g., "John Smith"
JzRemarks = all after last ";", e.g., "Who is the Boss?"

You can ignore Remarks, if want, and work as all others.
But Remarks is usefull to view and edit INIs like NotePad.

There are 3 principal procedures:

JzReadINI - Boolean Function : if true, returns JzValue, JzRemarks and JzValuINI.
JzWriteINI - Sub: Writes or Update JzValuINI (or separated JzValue + JzRemarks).
JzGetINI - is a combo: try read, if fail then create it. Needs default values
           into JzValue + JzRemarks or JzValuINI.

Obs: JzValue is the principal element to interface:
     JzValue = "" (empty), routines will try split JzValuINI into JzValue + JzRemarks.
     JzValue <> "", routines will compose JzValuINI from JzValue + JzRemarks.

======================================================================================
WELL, THAT'S ALL! I HOPE YOU LIKE IT!
(If you make fixes or enhances, please, send to me: jozew@globo.com  ... Thanks)

[Joze] JOZE Walter de Moura
from Rio de Janeiro, Brazil.

-oOo-oOo-oOo-







