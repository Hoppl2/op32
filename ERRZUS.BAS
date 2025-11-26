s$ = DIR$("*.frm")
t$ = "andi.xxx"
WHILE (s$ <> "")
  f% = FREEFILE
  OPEN s$ FOR INPUT AS #f%
  g% = FREEFILE
  OPEN t$ FOR OUTPUT AS #g%

  WHILE NOT EOF(f%)
    LINE INPUT #f%, h$
    PRINT #g%, h$

    IF (LEFT$(h$, 4) = "Call") THEN
      ind% = INSTR(h$, "DefErrFnc")
      IF (ind% > 0) THEN
        z$ = LEFT$(h$, ind% - 1) + "DefErrMod(DefErrModul)"
        PRINT #g%, z$
      END IF
    END IF
  WEND

  CLOSE #f%
  CLOSE #g%

  KILL s$
  NAME t$ AS s$

  s$ = DIR$
WEND

