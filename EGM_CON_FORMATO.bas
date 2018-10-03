


'==========================================| PERFIL | ========================================
Sub PERFIL()

Dim NOMBRE As String
Dim ERROR As String
Dim OK As String


ERROR = "LA MACRO SELECCIONADA NO ES VALIDA"
NOMBRE = "PERFIL"
OK = "OK"

If ActiveSheet.Name = NOMBRE Then

    Call CONTROLADOR(0101)
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B5").Select
    Selection.Cut
    Range("A6").Select
    ActiveSheet.Paste
    Range("B8").Select
    Selection.Cut
    Range("A9").Select
    ActiveSheet.Paste
    Range("B11").Select
    Selection.Cut
    Range("A12").Select
    ActiveSheet.Paste
    Range("B17").Select
    Selection.Cut
    Range("A18").Select
    ActiveSheet.Paste
    Range("B29").Select
    Selection.Cut
    Range("A30").Select
    ActiveSheet.Paste
    Range("B37").Select
    Selection.Cut
    Range("A38").Select
    ActiveSheet.Paste
    Range("B43").Select
    Selection.Cut
    Range("A44").Select
    ActiveSheet.Paste
    Range("B47").Select
    Selection.Cut
    Range("A48").Select
    ActiveSheet.Paste
    Range("B50").Select
    Selection.Cut
    Range("A51").Select
    ActiveSheet.Paste
    Range("B56").Select
    Selection.Cut
    Range("A57").Select
    ActiveSheet.Paste
    Range("B59").Select
    Selection.Cut
    Range("A60").Select
    ActiveSheet.Paste
    Range("B68").Select
    Selection.Cut
    Range("A69").Select
    ActiveSheet.Paste
    Rows("4:4").Select
    Selection.Delete Shift:=xlUp
    Selection.Delete Shift:=xlUp
    Rows("6:6").Select
    Selection.Delete Shift:=xlUp
    Rows("8:8").Select
    Selection.Delete Shift:=xlUp
    Rows("13:13").Select
    Selection.Delete Shift:=xlUp
    Rows("24:24").Select
    Selection.Delete Shift:=xlUp
    Rows("31:31").Select
    Selection.Delete Shift:=xlUp
    Rows("36:36").Select
    Selection.Delete Shift:=xlUp
    Rows("39:39").Select
    Selection.Delete Shift:=xlUp
    Range("41:41,47:47,50:50").Select
    Range("A50").Activate
    Selection.Delete Shift:=xlUp
    Rows("56:56").Select
    Selection.Delete Shift:=xlUp
    Range("B68").Select
    ActiveWindow.ScrollRow = 49
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 47
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 45
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 1
    Range("A4").Select
    Selection.Copy
    Range("A5").Select
    ActiveSheet.Paste
    Range("A6").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A7").Select
    ActiveSheet.Paste
    Range("A8").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A9:A12").Select
    ActiveSheet.Paste
    Range("A13").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A14:A23").Select
    ActiveSheet.Paste
    Range("A24").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A25:A30").Select
    ActiveSheet.Paste
    Range("A31").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A32:A35").Select
    ActiveSheet.Paste
    Range("A36").Select
    Columns("A:A").ColumnWidth = 35.63
    Application.CutCopyMode = False
    Selection.Copy
    Range("A37").Select
    ActiveSheet.Paste
    Range("A38").Select
    ActiveSheet.Paste
    Range("A39").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A40").Select
    ActiveSheet.Paste
    Range("A41").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A42:A45").Select
    ActiveSheet.Paste
    Range("A46").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A47").Select
    ActiveSheet.Paste
    Range("A48").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A49:A55,A41").Select
    Range("A41").Activate
    ActiveSheet.Paste
    Range("A56").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A57:A64").Select
    ActiveSheet.Paste
    Range("A56").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Rows("4:64").Select
    Application.CutCopyMode = False
    Rows("4:64").EntireRow.AutoFit
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("L:L").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    Columns("Q:Q").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("V:V").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AA:AA").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 21
    Columns("AF:AF").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AK:AK").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    Columns("AP:AP").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AU:AU").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    Columns("AZ:AZ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BE:BE").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 52
    Columns("BJ:BJ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BO:BO").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    Columns("BT:BT").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BY:BY").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    Columns("CD:CD").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CI:CI").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    Columns("CN:CN").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CW:CW").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 1
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D1").Select
    Selection.Cut
    Range("C4").Select
    ActiveSheet.Paste
    Range("C4").Select
    Selection.Copy
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("C64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("C5:C64").Select
    Range("C64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Range("I2").Select
    Selection.Cut
    Range("H4").Select
    ActiveSheet.Paste
    Columns("H:H").ColumnWidth = 14.13
    Selection.Copy
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("H64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("H5:H64").Select
    Range("H64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H4").Select
    Application.CutCopyMode = False
    Range("N2").Select
    Selection.Cut
    Range("M4").Select
    ActiveSheet.Paste
    Columns("M:M").ColumnWidth = 13.63
    Selection.Copy
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("M64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("M5:M64").Select
    Range("M64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("S2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("R4").Select
    ActiveSheet.Paste
    Columns("R:R").ColumnWidth = 23.88
    Columns("R:R").ColumnWidth = 37.5
    Columns("R:R").ColumnWidth = 39.88
    Selection.Copy
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("R64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("R5:R64").Select
    Range("R64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("X2").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("W4").Select
    ActiveSheet.Paste
    Columns("W:W").ColumnWidth = 21.75
    Application.CutCopyMode = False
    Selection.Copy
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("W64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("W5:W64").Select
    Range("W64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AJ2").Select
    Columns("AB:AB").ColumnWidth = 25
    Range("AC2").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AB4").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("AB64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AB5:AB64").Select
    Range("AB64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ1").Select
    Columns("AG:AG").ColumnWidth = 26
    Range("AH2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AG4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AG64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AG5:AG64").Select
    Range("AG64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AX2").Select
    Columns("AL:AL").ColumnWidth = 26.75
    Range("AM2").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AL4").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AL64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AL5:AL64").Select
    Range("AL64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BC1").Select
    Columns("AQ:AQ").ColumnWidth = 26
    Range("AR2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AQ4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AQ64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AQ5:AQ64").Select
    Range("AQ64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BJ1").Select
    Columns("AV:AV").ColumnWidth = 20.5
    Range("AW2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("AV5:AV64").Select
    ActiveSheet.Paste
    Range("BB3").Select
    Columns("BA:BA").ColumnWidth = 23.75
    Range("BB2").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BB2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BA4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("BA64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BA5:BA64").Select
    Range("BA64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Columns("BF:BF").ColumnWidth = 21.75
    Range("BG2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BF4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BF64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BF5:BF64").Select
    Range("BF64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BV4").Select
    Columns("BK:BK").ColumnWidth = 27.75
    Range("BL2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BK4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BK64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BK5:BK64").Select
    Range("BK64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ3").Select
    Columns("BP:BP").ColumnWidth = 28.5
    Range("BQ2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BP4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BP64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BP5:BP64").Select
    Range("BP64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CG1").Select
    Columns("BU:BU").ColumnWidth = 26.38
    Range("BV2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BU4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BU64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BU5:BU64").Select
    Range("BU64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CL1").Select
    Columns("BZ:BZ").ColumnWidth = 29
    Range("CA2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BZ4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BZ64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BZ5:BZ64").Select
    Range("BZ64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CR4").Select
    Columns("CE:CE").ColumnWidth = 27
    Range("CF2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("CE4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("BZ5").Select
    Selection.End(xlDown).Select
    Range("CE64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CE5:CE64").Select
    Range("CE64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CN3").Select
    Columns("CJ:CJ").ColumnWidth = 17.38
    Range("CK2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("CJ4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CJ64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CJ5:CJ64").Select
    Range("CJ64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CW4").Select
    Columns("CO:CO").ColumnWidth = 23.25
    Range("CP2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("CO4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CO64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CO5:CO64").Select
    Range("CO64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Columns("CT:CT").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CU5").Select
    Columns("CT:CT").ColumnWidth = 16.75
    Range("CU2").Select
    Selection.Cut
    Range("CT4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CT64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CT5:CT64").Select
    Range("CT64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CZ2").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CY4").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CY64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CY5:CY64").Select
    Range("CY64").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CZ5").Select
    Application.CutCopyMode = False
    Range("CM4").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range("BZ4").Select
    Selection.End(xlToRight).Select
    Range("CY4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CT65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CO65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CJ65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CE65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("BZ65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BU65").Select
    ActiveSheet.Paste
    Range("BU64").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BP65").Select
    ActiveSheet.Paste
    Range("BP64").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BK65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BF65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BA65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("AV65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AQ65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AL65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AG65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AB65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("W65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("R65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("M65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("H65").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlDown).Select
    Range("H4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("C65").Select
    ActiveSheet.Paste
    Range("C64").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range("B64").Select
    Columns("B:B").ColumnWidth = 36.25
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A4:B64").Select
    Range("B64").Activate
    Selection.Copy
    Range("C64").Select
    Selection.End(xlDown).Select
    Range("A1284:B1284").Select
    Range("B1284").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A65:B1284").Select
    Range("B1284").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Columns("H:DC").Select
    Selection.Delete Shift:=xlToLeft
    Rows("1:1").Select
    Range("B1").Activate
    Selection.Delete Shift:=xlUp
    Rows("2:2").Select
    Range("B2").Activate
    Selection.Delete Shift:=xlUp
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Target"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Respuesta"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Modalidades"
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Variable"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "PERFIL"
    Range("A2").Select
    Selection.Copy
    Range("B2").Select
    Selection.End(xlDown).Select
    Range("A1282").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A3:A1282").Select
    Range("A1282").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Range("A2:H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearFormats
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    Range("C1").Select
    Selection.Copy
    Range("B1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A5").Select
    Selection.End(xlDown).Select
    Range("A1282").Select
    Selection.End(xlUp).Select

    CLEARCONTROLER(0101)
    

    Else
    MsgBox ERROR
End If
End Sub

'==========================================| FIN PERFIL | ========================================


'==========================================| EQUIPAMIENTO | ========================================


Sub EQUIPAMIENTO()
Dim NOMBRE As String
Dim ERROR As String
Dim OK As String


ERROR = "LA MACRO SELECCIONADA NO ES VALIDA"
NOMBRE = "EQUIPAMIENTO"



If ActiveSheet.Name = NOMBRE Then

'
' Macro3 Macro
'
' Acceso directo: CTRL+i
'
    CONTROLADOR(0202)
    Range("H11").Select
    ActiveWindow.SmallScroll Down:=-24
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.ColumnWidth = 37.25
    Range("B5").Select
    Selection.Cut
    Range("A6").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=12
    Range("B26").Select
    Selection.Cut
    Range("A27").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=12
    Range("B33").Select
    Selection.Cut
    Range("A34").Select
    ActiveSheet.Paste
    Range("B39").Select
    Selection.Cut
    Range("A40").Select
    ActiveSheet.Paste
    Range("B42").Select
    Selection.Cut
    Range("A43").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=9
    Range("B47").Select
    Selection.Cut
    Range("A48").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=-51
    Rows("4:5").Select
    Selection.Delete Shift:=xlUp
    Range("B6").Select
    ActiveWindow.SmallScroll Down:=18
    Rows("24:24").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=9
    Range("36:36,39:39,44:44").Select
    Range("A44").Activate
    Selection.Delete Shift:=xlUp
    Range("A33").Select
    ActiveWindow.SmallScroll Down:=-45
    Range("A4").Select
    Rows("4:4").RowHeight = 22.5
    Range("A4").Select
    Selection.Copy
    Range("A5").Select
    ActiveWindow.SmallScroll Down:=6
    Range("A5:A23").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=15
    Range("A24").Select
    Application.CutCopyMode = False
    Selection.Cut
    Application.CutCopyMode = False
    Selection.Copy
    Range("A25:A30").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=6
    Range("A31").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A32:A35").Select
    ActiveSheet.Paste
    Range("A36").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A37").Select
    ActiveSheet.Paste
    Range("A38").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A39:A41").Select
    ActiveSheet.Paste
    Range("A42").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A43").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.End(xlUp).Select
    ActiveWindow.SmallScroll Down:=-27
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D1").Select
    Selection.Cut
    Range("C4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 2
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("I2").Select
    Selection.Cut
    Range("H4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("N2").Select
    Selection.Cut
    Range("M4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    Columns("R:R").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.ColumnWidth = 18.13
    Range("S2").Select
    Selection.Cut
    Range("R4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    Columns("W:W").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("X2").Select
    Selection.Cut
    Range("W4").Select
    ActiveSheet.Paste
    Columns("AB:AB").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AC2").Select
    Selection.Cut
    Range("AB4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    Columns("AG:AG").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AH2").Select
    Selection.Cut
    Range("AG4").Select
    ActiveSheet.Paste
    Columns("AL:AL").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AM2").Select
    Selection.Cut
    Range("AL4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 36
    Columns("AQ:AQ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AR2").Select
    Selection.Cut
    Range("AQ4").Select
    ActiveSheet.Paste
    Columns("AV:AV").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AW2").Select
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 44
    Columns("BA:BA").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BB2").Select
    Selection.Cut
    Range("BA4").Select
    ActiveSheet.Paste
    Columns("BF:BF").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BG2").Select
    Selection.Cut
    Range("BF4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    Columns("BK:BK").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BL2").Select
    Selection.Cut
    Range("BK4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    Columns("BP:BP").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BQ2").Select
    Selection.Cut
    Range("BP4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    Columns("BU:BU").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BV2").Select
    Selection.Cut
    Range("BU4").Select
    ActiveSheet.Paste
    Columns("BZ:BZ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CA2").Select
    Selection.Cut
    Range("BZ4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    Columns("CE:CE").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CF2").Select
    Selection.Cut
    Range("CE4").Select
    ActiveSheet.Paste
    Columns("CJ:CJ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CK2").Select
    Selection.Cut
    Range("CJ4").Select
    ActiveSheet.Paste
    Columns("CO:CO").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CP2").Select
    Selection.Cut
    Range("CO4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    Columns("CT:CT").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CU2").Select
    Selection.Cut
    Range("CT4").Select
    ActiveSheet.Paste
    Columns("CY:CY").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CZ2").Select
    Selection.Cut
    Range("CY4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("C4").Select
    Selection.Copy
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("C50").Select
    ActiveWindow.SmallScroll Down:=-18
    Application.CutCopyMode = False
    ActiveWindow.SmallScroll Down:=-15
    Rows("30:30").Select
    Selection.Delete Shift:=xlUp
    Range("A37").Select
    ActiveWindow.SmallScroll Down:=-6
    Rows("24:29").Select
    Selection.RowHeight = 15.75
    Range("A31").Select
    ActiveWindow.SmallScroll Down:=-36
    Range("C4").Select
    Selection.Copy
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("C42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("C5:C42").Select
    Range("C42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G4").Select
    Selection.End(xlDown).Select
    Range("H42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("H5:H42").Select
    Range("H42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("L4").Select
    Selection.End(xlDown).Select
    Range("M42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("M5:M42").Select
    Range("M42").Activate
    ActiveSheet.Paste
    Range("C4").Select
    Application.CutCopyMode = False
    Range("R4").Select
    Selection.Copy
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("R42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("R5:R42").Select
    Range("R42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("R4,R6").Select
    Range("R6").Activate
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("W42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("W5:W42").Select
    Range("W42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("AB42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AB5:AB42").Select
    Range("AB42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AG42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AG5:AG42").Select
    Range("AG42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AL42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AL5:AL42").Select
    Range("AL42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AQ42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AQ5:AQ42").Select
    Range("AQ42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AV42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AV5:AV42").Select
    Range("AV42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("BA42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BA5:BA42").Select
    Range("BA42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BF42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BF5:BF42").Select
    Range("BF42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BF5").Select
    Selection.End(xlDown).Select
    Range("BK42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BK5:BK42").Select
    Range("BK42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BP42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BP5:BP42").Select
    Range("BP42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BU42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BU5:BU42").Select
    Range("BU42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BZ42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BZ5:BZ42").Select
    Range("BZ42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("CE42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CE5:CE42").Select
    Range("CE42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CJ42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CJ5:CJ42").Select
    Range("CJ42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CO42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CO5:CO42").Select
    Range("CO42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CT42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CT5:CT42").Select
    Range("CT42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CY4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CY42").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CY5:CY42").Select
    Range("CY42").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CY4").Select
    Application.CutCopyMode = False
    ActiveWindow.ScrollColumn = 90
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Range("D1:G1").Select
    Selection.Cut
    Range("D2").Select
    ActiveSheet.Paste
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Range("I1").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 90
    Range("CY2:DC2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CT2").Select
    Selection.End(xlDown).Select
    Range("CT41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CO2").Select
    Selection.End(xlDown).Select
    Range("CO41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CJ2").Select
    Selection.End(xlDown).Select
    Range("CJ41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE2").Select
    Selection.End(xlDown).Select
    Range("CE41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BZ2").Select
    Selection.End(xlDown).Select
    Range("BZ41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BU2").Select
    Selection.End(xlDown).Select
    Range("BU41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BP2").Select
    Selection.End(xlDown).Select
    Range("BP41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BK2").Select
    Selection.End(xlDown).Select
    Range("BK41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BF2").Select
    Selection.End(xlDown).Select
    Range("BF41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BA2").Select
    Selection.End(xlDown).Select
    Range("BA41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AV2").Select
    Selection.End(xlDown).Select
    Range("AV41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AQ2").Select
    Selection.End(xlDown).Select
    Range("AQ41").Select
    ActiveSheet.Paste
    Range("AQ40").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AL2").Select
    Selection.End(xlDown).Select
    Range("AL41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG2").Select
    Selection.End(xlDown).Select
    Range("AG41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AB2").Select
    Selection.End(xlDown).Select
    Range("AB41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W2").Select
    Selection.End(xlDown).Select
    Range("W41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("R2").Select
    Selection.End(xlDown).Select
    Range("R41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("M2").Select
    Selection.End(xlDown).Select
    Range("M41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H2").Select
    Selection.End(xlDown).Select
    Range("H41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("C2").Select
    Selection.End(xlDown).Select
    Range("C41").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("C820").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A2:B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("C2").Select
    Selection.End(xlDown).Select
    Range("A820:B820").Select
    Range("B820").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A41:B820").Select
    Range("B820").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Columns("H:H").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 90
    Columns("H:DC").Select
    Selection.Delete Shift:=xlToLeft
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Target"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Respuesta"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Modalidades"
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Variables"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "EQUIPAMIENTO"
    Range("A2").Select
    Selection.Copy
    Range("B2").Select
    Selection.End(xlDown).Select
    Range("A820").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A3:A820").Select
    Range("A820").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Columns("A:A").ColumnWidth = 23.13
    Range("A2:H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearFormats
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    Range("C1").Select
    Selection.Copy
    Range("B1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A2").Select
    Rows("2:2").RowHeight = 17
    Columns("C:C").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Selection.End(xlDown).Select
    Range("A820").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select


    CLEARCONTROLER(0202)

    Else
    MsgBox ERROR
End If


End Sub




'==========================================| FIN EQUIPAMIENTO | ========================================

'==========================================| ESTILODEVIDA | ========================================

Sub ESTILOSDEVIDA()

Dim NOMBRE As String
Dim ERROR As String
Dim OK As String


ERROR = "LA MACRO SELECCIONADA NO ES VALIDA"
NOMBRE = "ESTILOS DE VIDA"
OK = "OK"


If ActiveSheet.Name = NOMBRE Then
   


'
' Macro1 Macro
'
' Acceso directo: CTRL+y
'
    CONTROLADOR(0303)
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.ColumnWidth = 42.38
    Range("B5").Select
    Selection.Cut
    Range("A6").Select
    ActiveSheet.Paste
    Range("B13").Select
    Selection.Cut
    Range("A14").Select
    ActiveSheet.Paste
    Range("B95").Select
    Selection.Cut
    Range("A96").Select
    ActiveSheet.Paste
    Range("B140").Select
    Selection.Cut
    Range("A141").Select
    ActiveSheet.Paste
    Range("B164").Select
    Selection.Cut
    Range("A165").Select
    ActiveSheet.Paste
    Range("B177").Select
    Selection.Cut
    Range("A178").Select
    ActiveSheet.Paste
    Range("B255").Select
    Selection.Cut
    Range("A256").Select
    ActiveSheet.Paste
    Range("B268").Select
    Selection.Cut
    Range("A269").Select
    ActiveSheet.Paste
    Range("C272").Select
    Selection.End(xlDown).Select
    Range("B290").Select
    Selection.Cut
    Range("A290").Select
    ActiveSheet.Paste
    Range("C290").Select
    Selection.End(xlDown).Select
    Range("C292").Select
    Selection.End(xlDown).Select
    Range("B324").Select
    Selection.Cut
    Range("A325").Select
    ActiveSheet.Paste
    Range("C325").Select
    Selection.End(xlDown).Select
    Range("B331").Select
    Selection.Cut
    Range("A332").Select
    ActiveSheet.Paste
    Range("C332").Select
    Selection.End(xlDown).Select
    ActiveWindow.SmallScroll Down:=-42
    ActiveWindow.ScrollRow = 299
    ActiveWindow.ScrollRow = 298
    ActiveWindow.ScrollRow = 297
    ActiveWindow.ScrollRow = 296
    ActiveWindow.ScrollRow = 294
    ActiveWindow.ScrollRow = 292
    ActiveWindow.ScrollRow = 289
    ActiveWindow.ScrollRow = 284
    ActiveWindow.ScrollRow = 278
    ActiveWindow.ScrollRow = 257
    ActiveWindow.ScrollRow = 225
    ActiveWindow.ScrollRow = 164
    ActiveWindow.ScrollRow = 142
    ActiveWindow.ScrollRow = 110
    ActiveWindow.ScrollRow = 97
    ActiveWindow.ScrollRow = 71
    ActiveWindow.ScrollRow = 61
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 45
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 42
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 32
    ActiveWindow.ScrollRow = 29
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    Rows("4:5").Select
    Selection.Delete Shift:=xlUp
    Rows("11:11").Select
    ActiveWindow.SmallScroll Down:=72
    Range("11:11,93:93").Select
    Range("A93").Activate
    ActiveWindow.SmallScroll Down:=54
    Range("11:11,93:93,138:138").Select
    Range("A138").Activate
    ActiveWindow.SmallScroll Down:=21
    Range("11:11,93:93,138:138,162:162").Select
    Range("A162").Activate
    ActiveWindow.SmallScroll Down:=21
    Range("11:11,93:93,138:138,162:162,175:175").Select
    Range("A175").Activate
    ActiveWindow.SmallScroll Down:=75
    Range("11:11,93:93,138:138,162:162,175:175,253:253,266:266").Select
    Range("A266").Activate
    ActiveWindow.SmallScroll Down:=39
    Range("11:11,93:93,138:138,162:162,175:175,253:253,266:266,287:287").Select
    Range("A287").Activate
    ActiveWindow.SmallScroll Down:=-9
    Range("A288").Select
    Selection.Cut Destination:=Range("A289")
    Rows("288:288").Select
    Selection.Delete Shift:=xlUp
    Range("A283").Select
    ActiveWindow.ScrollRow = 273
    ActiveWindow.ScrollRow = 272
    ActiveWindow.ScrollRow = 271
    ActiveWindow.ScrollRow = 270
    ActiveWindow.ScrollRow = 268
    ActiveWindow.ScrollRow = 267
    ActiveWindow.ScrollRow = 265
    ActiveWindow.ScrollRow = 263
    ActiveWindow.ScrollRow = 261
    ActiveWindow.ScrollRow = 259
    ActiveWindow.ScrollRow = 256
    ActiveWindow.ScrollRow = 252
    ActiveWindow.ScrollRow = 246
    ActiveWindow.ScrollRow = 237
    ActiveWindow.ScrollRow = 201
    ActiveWindow.ScrollRow = 173
    ActiveWindow.ScrollRow = 115
    ActiveWindow.ScrollRow = 90
    ActiveWindow.ScrollRow = 58
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 36
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 1
    Rows("11:11").Select
    Selection.Delete Shift:=xlUp
    Range("A17").Select
    ActiveWindow.SmallScroll Down:=69
    Rows("92:92").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=51
    Rows("136:136").Select
    Selection.Delete Shift:=xlUp
    Range("A141").Select
    ActiveWindow.SmallScroll Down:=30
    Rows("159:159").Select
    Selection.Delete Shift:=xlUp
    Rows("171:171").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=87
    Rows("248:248").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=18
    Rows("260:260").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=48
    Rows("314:314").Select
    Selection.Delete Shift:=xlUp
    Range("A311").Select
    ActiveWindow.SmallScroll Down:=3
    Rows("320:320").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=-6
    ActiveWindow.ScrollRow = 300
    ActiveWindow.ScrollRow = 299
    ActiveWindow.ScrollRow = 296
    ActiveWindow.ScrollRow = 294
    ActiveWindow.ScrollRow = 291
    ActiveWindow.ScrollRow = 287
    ActiveWindow.ScrollRow = 279
    ActiveWindow.ScrollRow = 274
    ActiveWindow.ScrollRow = 259
    ActiveWindow.ScrollRow = 233
    ActiveWindow.ScrollRow = 175
    ActiveWindow.ScrollRow = 148
    ActiveWindow.ScrollRow = 108
    ActiveWindow.ScrollRow = 96
    ActiveWindow.ScrollRow = 84
    ActiveWindow.ScrollRow = 81
    ActiveWindow.ScrollRow = 72
    ActiveWindow.ScrollRow = 65
    ActiveWindow.ScrollRow = 63
    ActiveWindow.ScrollRow = 62
    ActiveWindow.ScrollRow = 60
    ActiveWindow.ScrollRow = 59
    ActiveWindow.ScrollRow = 58
    ActiveWindow.ScrollRow = 57
    ActiveWindow.ScrollRow = 56
    ActiveWindow.ScrollRow = 55
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 53
    ActiveWindow.ScrollRow = 52
    ActiveWindow.ScrollRow = 51
    ActiveWindow.ScrollRow = 49
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 47
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 45
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 42
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 37
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 34
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 30
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 26
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 1
    Range("A4").Select
    Selection.Copy
    Range("A5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A5:A10").Select
    ActiveSheet.Paste
    Range("A11").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A11").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A12").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A12:A91").Select
    ActiveSheet.Paste
    Range("A13").Select
    Selection.End(xlDown).Select
    Range("A92").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A93").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A93:A135").Select
    ActiveSheet.Paste
    Range("A95").Select
    Selection.End(xlDown).Select
    Range("A136").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A137").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A137:A158").Select
    ActiveSheet.Paste
    Range("A140").Select
    Selection.End(xlDown).Select
    Range("A159").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A160").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A160:A170").Select
    ActiveSheet.Paste
    Range("A171").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A172").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A172:A247").Select
    ActiveSheet.Paste
    Range("A175").Select
    Selection.End(xlDown).Select
    Range("A248").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A249").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A249:A259").Select
    ActiveSheet.Paste
    Range("A260").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A261").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A261:A280").Select
    ActiveSheet.Paste
    Range("A265").Select
    Selection.End(xlDown).Select
    Range("A281").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.End(xlDown).Select
    Range("A313").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A282:A313").Select
    Range("A313").Activate
    ActiveSheet.Paste
    Range("A314").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A315:A319").Select
    ActiveSheet.Paste
    Range("A320").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A321").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A321:A1048574").Select
    Selection.End(xlUp).Select
    Range("B321").Select
    Selection.End(xlDown).Select
    Range("A347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A321:A347").Select
    Range("A347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D1").Select
    Selection.Cut
    Range("C4").Select
    ActiveSheet.Paste
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("I2").Select
    Selection.Cut
    Range("H4").Select
    ActiveSheet.Paste
    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("N2").Select
    Selection.Cut
    Range("M4").Select
    ActiveSheet.Paste
    Columns("R:R").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("S2").Select
    Selection.Cut
    Range("R4").Select
    ActiveSheet.Paste
    Columns("W:W").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("X2").Select
    Selection.Cut
    Range("W4").Select
    ActiveSheet.Paste
    Columns("AB:AB").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AC2").Select
    Selection.Cut
    Range("AB4").Select
    ActiveSheet.Paste
    Columns("AG:AG").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AH2").Select
    Selection.Cut
    Range("AG4").Select
    ActiveSheet.Paste
    Columns("AL:AL").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AM2").Select
    Selection.Cut
    Range("AL4").Select
    ActiveSheet.Paste
    Columns("AQ:AQ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AR2").Select
    Selection.Cut
    Range("AQ4").Select
    ActiveSheet.Paste
    Columns("AV:AV").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("AW2").Select
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    Columns("BA:BA").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BB2").Select
    Selection.Cut
    Range("BA4").Select
    ActiveSheet.Paste
    Columns("BF:BF").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BG2").Select
    Selection.Cut
    Range("BF4").Select
    ActiveSheet.Paste
    Columns("BK:BK").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BL2").Select
    Selection.Cut
    Range("BK4").Select
    ActiveSheet.Paste
    Columns("BP:BP").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BQ2").Select
    Selection.Cut
    Range("BP4").Select
    ActiveSheet.Paste
    Columns("BU:BU").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("BV2").Select
    Selection.Cut
    Range("BU4").Select
    ActiveSheet.Paste
    Columns("BZ:BZ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CE:CE").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CF2").Select
    Selection.Cut
    Range("CE4").Select
    ActiveSheet.Paste
    Range("CA2").Select
    Selection.Cut
    Range("BZ4").Select
    ActiveSheet.Paste
    Columns("CJ:CJ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CO:CO").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CT:CT").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CK2").Select
    Selection.Cut
    Range("CJ4").Select
    ActiveSheet.Paste
    Range("CP2").Select
    Selection.Cut
    Range("CO4").Select
    ActiveSheet.Paste
    Range("CU2").Select
    Selection.Cut
    Range("CT4").Select
    ActiveSheet.Paste
    Columns("CY:CY").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("CZ2").Select
    Selection.Cut
    Range("CY4").Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("C4").Select
    Selection.Copy
    Application.CutCopyMode = False
    Range("C4").Select
    Selection.Copy
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("C347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("C5:C347").Select
    Range("C347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("H347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("H5:H347").Select
    Range("H347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("M347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("M5:M347").Select
    Range("M347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("R347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("R5:R347").Select
    Range("R347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("W347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("W5:W347").Select
    Range("W347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("AB347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AB5:AB347").Select
    Range("AB347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AG347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AG5:AG347").Select
    Range("AG347").Activate
    ActiveSheet.Paste
    Range("AG346").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AL347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AL5:AL347").Select
    Range("AL347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AQ347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AQ5:AQ347").Select
    Range("AQ347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AV347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AV5:AV347").Select
    Range("AV347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("BA347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BA5:BA347").Select
    Range("BA347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BF347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BF5:BF347").Select
    Range("BF347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BK347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BK5:BK347").Select
    Range("BK347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BP347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BP5:BP347").Select
    Range("BP347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BU347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BU5:BU347").Select
    Range("BU347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BZ347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BZ5:BZ347").Select
    Range("BZ347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("CE347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CE5:CE347").Select
    Range("CE347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CJ347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CJ5:CJ347").Select
    Range("CJ347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CO347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CO5:CO347").Select
    Range("CO347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CT347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CT5:CT347").Select
    Range("CT347").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CY4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CY347").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CY5:CY347").Select
    Range("CY347").Activate
    ActiveSheet.Paste
    Range("CY346").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("DB4").Select
    Application.CutCopyMode = False
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range("D2:G2").Select
    Selection.Cut Destination:=Range("D3:G3")
    Rows("1:2").Select
    Selection.Delete Shift:=xlUp
    Range("E4").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 90
    Range("CY1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=-150
    ActiveWindow.ScrollRow = 188
    ActiveWindow.ScrollRow = 186
    ActiveWindow.ScrollRow = 178
    ActiveWindow.ScrollRow = 173
    ActiveWindow.ScrollRow = 128
    ActiveWindow.ScrollRow = 91
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    Range("CY2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CT2").Select
    Selection.End(xlDown).Select
    Range("CT346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CO2").Select
    Selection.End(xlDown).Select
    Range("CO346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CJ2").Select
    Selection.End(xlDown).Select
    Range("CJ346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE2").Select
    Selection.End(xlDown).Select
    Range("CE346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BZ2").Select
    Selection.End(xlDown).Select
    Range("BZ346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BU2").Select
    Selection.End(xlDown).Select
    Range("BU346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BP2").Select
    Selection.End(xlDown).Select
    Range("BP346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BK2").Select
    Selection.End(xlDown).Select
    Range("BK346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BF2").Select
    Selection.End(xlDown).Select
    Range("BF346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BA2").Select
    Selection.End(xlDown).Select
    Range("BA346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AV2").Select
    Selection.End(xlDown).Select
    Range("AV346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AQ2").Select
    Selection.End(xlDown).Select
    Range("AQ346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AL2").Select
    Selection.End(xlDown).Select
    Range("AL346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG2").Select
    Selection.End(xlDown).Select
    Range("AG346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AB2").Select
    Selection.End(xlDown).Select
    Range("AB346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W2").Select
    Selection.End(xlDown).Select
    Range("W346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("R2").Select
    Selection.End(xlDown).Select
    Range("R346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("M2").Select
    Selection.End(xlDown).Select
    Range("M346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H2").Select
    Selection.End(xlDown).Select
    Range("H346").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("C2").Select
    Selection.End(xlDown).Select
    Range("C346").Select
    ActiveSheet.Paste
    Range("C345").Select
    Selection.End(xlUp).Select
    Range("C1").Select
    Selection.End(xlUp).Select
    Range("A94").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A2:B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("C2").Select
    Selection.End(xlDown).Select
    Range("A7225:B7225").Select
    Range("B7225").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A346:B7225").Select
    Range("B7225").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("C2").Select
    Application.CutCopyMode = False
    Columns("H:H").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 90
    Columns("H:DC").Select
    Selection.Delete Shift:=xlToLeft
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Target"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Respuestas"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Modalidades"
    Range("B1").Select
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A2:G2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearFormats
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.499984740745262
        .Weight = xlThin
    End With
    Range("A17").Select
    ActiveWindow.SmallScroll Down:=-21
    Range("A1").Select
    Selection.AutoFilter
    Range("A16").Select
    Selection.End(xlDown).Select
    Range("A7225").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select

    CLEARCONTROLER(0303)

    Else
    MsgBox ERROR
End If
End Sub

'==========================================| FIN ESTILODEVIDA | ========================================

'==========================================| CONSUMOINDIVIDUO | ========================================



Sub CONSUMOINDIVIDUO()

Dim NOMBRE As String
Dim ERROR As String
Dim OK As String


ERROR = "LA MACRO SELECCIONADA NO ES VALIDA"
NOMBRE = "CONSUMO INDIVIDUO"



If ActiveSheet.Name = NOMBRE Then
   


'
' CONSUMOINDIVIDUOREP Macro
'

'

    CONTROLADOR(0404)    
    Range("A4").Select
    Selection.EntireColumn.Insert
    Range("B4").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=9
    Selection.Copy
    ActiveWindow.SmallScroll Down:=-63
    Range("A4").Select
    ActiveSheet.Paste
    Range("A5").Select
    Application.CutCopyMode = False
    Range("A6:A10").Select
    Selection.ClearContents
    Range("A12:A15").Select
    Selection.ClearContents
    Range("C16").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A17:A32").Select
    Range("A32").Activate
    Selection.ClearContents
    Range("C32").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A43").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A34:A43").Select
    Range("A43").Activate
    Selection.ClearContents
    Range("C43").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A54").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A45:A54").Select
    Range("A54").Activate
    Selection.ClearContents
    Range("C54").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A64").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A56:A64").Select
    Range("A64").Activate
    Selection.ClearContents
    Range("C64").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A69").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A66:A69").Select
    Range("A69").Activate
    Selection.ClearContents
    Range("C69").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A73").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A71:A73").Select
    Range("A73").Activate
    Selection.ClearContents
    Range("C73").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A77").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A75:A77").Select
    Range("A77").Activate
    Selection.ClearContents
    Range("C77").Select
    Selection.End(xlDown).Select
    Range("C1048576").Select
    Selection.End(xlUp).Select
    Range("C44").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A5").Select
    Selection.AutoFill Destination:=Range("A5:A10")
    Range("A5:A10").Select
    Range("A11").Select
    Selection.AutoFill Destination:=Range("A11:A15")
    Range("A11:A15").Select
    Range("A16").Select
    Selection.AutoFill Destination:=Range("A16:A32")
    Range("A16:A32").Select
    ActiveWindow.SmallScroll Down:=18
    Range("A33").Select
    Selection.AutoFill Destination:=Range("A33:A43")
    Range("A33:A43").Select
    ActiveWindow.SmallScroll Down:=21
    Range("A44").Select
    Selection.AutoFill Destination:=Range("A44:A54")
    Range("A44:A54").Select
    ActiveWindow.SmallScroll Down:=9
    Range("A55").Select
    Selection.AutoFill Destination:=Range("A55:A64")
    Range("A55:A64").Select
    ActiveWindow.SmallScroll Down:=9
    Range("A65").Select
    Selection.AutoFill Destination:=Range("A65:A69")
    Range("A65:A69").Select
    ActiveWindow.SmallScroll Down:=9
    Range("A70").Select
    Selection.AutoFill Destination:=Range("A70:A73")
    Range("A70:A73").Select
    Range("A74").Select
    Selection.AutoFill Destination:=Range("A74:A77")
    Range("A74:A77").Select
    Range("C72").Select
    ActiveWindow.SmallScroll Down:=-105
    Rows("5:5").Select
    Selection.Delete Shift:=xlUp
    Rows("4:4").Select
    Selection.Delete Shift:=xlUp
    Rows("9:9").Select
    Selection.Delete Shift:=xlUp
    Rows("13:13").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=15
    Rows("29:29").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=6
    Range("C30").Select
    Selection.End(xlDown).Select
    Rows("39:39").Select
    Selection.Delete Shift:=xlUp
    Range("C39").Select
    Selection.End(xlDown).Select
    ActiveWindow.SmallScroll Down:=15
    Rows("49:49").Select
    Selection.Delete Shift:=xlUp
    Rows("58:58").Select
    Selection.Delete Shift:=xlUp
    Range("D58").Select
    Selection.End(xlDown).Select
    ActiveWindow.SmallScroll Down:=6
    Rows("62:62").Select
    Selection.Delete Shift:=xlUp
    Rows("65:65").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=6
    Range("D67").Select
    Selection.End(xlUp).Select
    ActiveWindow.SmallScroll Down:=-48
    Range("C2:F2").Select
    Selection.Cut
    Range("C3").Select
    ActiveSheet.Paste
    Range("C1").Select
    Selection.Cut
    Range("B1").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("A3").Select
    ActiveSheet.Paste
    Range("C3").Select
    Selection.EntireColumn.Insert
    Range("H3").Select
    Selection.End(xlUp).Select
    Range("H3").Select
    Selection.EntireColumn.Insert
    Range("M3").Select
    Selection.EntireColumn.Insert
    Range("R3").Select
    Selection.EntireColumn.Insert
    Range("W3").Select
    Selection.EntireColumn.Insert
    Range("AB3").Select
    Selection.EntireColumn.Insert
    Range("AG3").Select
    Selection.EntireColumn.Insert
    Range("AL3").Select
    Selection.EntireColumn.Insert
    Range("AQ3").Select
    Selection.EntireColumn.Insert
    Range("AV3").Select
    Selection.EntireColumn.Insert
    Range("BA3").Select
    Selection.EntireColumn.Insert
    Range("BF3").Select
    Selection.EntireColumn.Insert
    Range("BK3").Select
    Selection.EntireColumn.Insert
    Range("BP3").Select
    Selection.EntireColumn.Insert
    Range("BU3").Select
    Selection.EntireColumn.Insert
    Range("BZ3").Select
    Selection.EntireColumn.Insert
    Range("CE3").Select
    Selection.EntireColumn.Insert
    Range("CJ3").Select
    Selection.EntireColumn.Insert
    Range("CO3").Select
    Selection.EntireColumn.Insert
    Range("CT3").Select
    Selection.EntireColumn.Insert
    Range("CY3").Select
    Selection.EntireColumn.Insert
    Range("CZ2").Select
    Selection.Cut
    Range("CY4").Select
    ActiveSheet.Paste
    Range("CY4").Select
    Selection.AutoFill Destination:=Range("CY4:CY67")
    Range("CY4:CY67").Select
    Range("CU2").Select
    Selection.Cut
    Range("CT4").Select
    ActiveSheet.Paste
    Range("CT4").Select
    Selection.AutoFill Destination:=Range("CT4:CT67")
    Range("CT4:CT67").Select
    Range("CP2").Select
    Selection.Cut
    Range("CO4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("CO4:CO67")
    Range("CO4:CO67").Select
    Range("CK2").Select
    Selection.Cut
    Range("CJ4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("CJ4:CJ67")
    Range("CJ4:CJ67").Select
    Range("CF2").Select
    Selection.Cut
    Range("CE4").Select
    ActiveSheet.Paste
    Range("CF4").Select
    Selection.End(xlDown).Select
    Range("CE67").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CE67").Select
    Selection.End(xlUp).Select
    Selection.Copy
    Range("CF4").Select
    Selection.End(xlDown).Select
    Range("CE67").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CA67").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BZ4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("CA4").Select
    Selection.End(xlDown).Select
    Range("BZ67").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BV67").Select
    Selection.End(xlUp).Select
    Range("BV2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BU4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("BV4").Select
    Selection.End(xlDown).Select
    Range("BU67").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BQ66").Select
    Selection.End(xlUp).Select
    Range("BQ2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BP4").Select
    ActiveSheet.Paste
    Range("BQ4").Select
    Selection.End(xlDown).Select
    Range("BQ67").Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Selection.Copy
    Range("BQ4").Select
    Selection.End(xlDown).Select
    Range("BP67").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BM66").Select
    Selection.End(xlUp).Select
    Range("BL2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BK4").Select
    ActiveSheet.Paste
    Range("BK4").Select
    Selection.Copy
    Range("BL4").Select
    Selection.End(xlDown).Select
    Range("BK67").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BF66").Select
    Selection.End(xlUp).Select
    Range("BG2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BF4").Select
    ActiveSheet.Paste
    Range("BG4").Select
    Selection.End(xlDown).Select
    Range("BF67").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BF66").Select
    Selection.End(xlUp).Select
    Selection.Copy
    Range("BG4").Select
    Selection.End(xlDown).Select
    Range("BF67").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BB67").Select
    Selection.End(xlUp).Select
    Range("BB2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BA4").Select
    ActiveSheet.Paste
    Range("BA4").Select
    Selection.End(xlToRight).Select
    Selection.End(xlToLeft).Select
    Range("AW2").Select
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    Range("BA4").Select
    Selection.AutoFill Destination:=Range("BA4:BA67")
    Range("BA4:BA67").Select
    Range("AV4").Select
    Selection.Copy
    Range("AW4").Select
    Selection.End(xlDown).Select
    Range("AV67").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AS67").Select
    Selection.End(xlUp).Select
    Range("AR2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AQ4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("AQ4:AQ67")
    Range("AQ4:AQ67").Select
    Range("AM2").Select
    Selection.Cut
    Range("AL4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("AM3").Select
    Selection.End(xlDown).Select
    Range("AL67").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AK66").Select
    Selection.End(xlToLeft).Select
    Range("AH65").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AG4").Select
    ActiveSheet.Paste
    Range("AG4").Select
    Selection.Copy
    Range("AH4").Select
    Selection.End(xlDown).Select
    Range("AG67").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AC67").Select
    Selection.End(xlUp).Select
    Range("AC2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AB4").Select
    ActiveSheet.Paste
    Range("AA4").Select
    Selection.End(xlToRight).Select
    Range("DB4").Select
    Selection.End(xlToLeft).Select
    Range("AB4").Select
    Selection.Copy
    Selection.End(xlToRight).Select
    Selection.End(xlToLeft).Select
    Range("AC4").Select
    Selection.End(xlDown).Select
    Range("AB67").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AB65").Select
    Selection.End(xlToLeft).Select
    Range("V64").Select
    Selection.End(xlUp).Select
    Range("X27").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("X2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("W4").Select
    ActiveSheet.Paste
    Range("W4").Select
    Selection.Copy
    Range("X4").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range("W67").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("R67").Select
    Selection.End(xlUp).Select
    Range("S2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("R4").Select
    ActiveSheet.Paste
    Range("R4").Select
    Selection.Copy
    Range("S4").Select
    Selection.End(xlDown).Select
    Range("R67").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("M67").Select
    Selection.End(xlUp).Select
    Range("N2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("M4").Select
    ActiveSheet.Paste
    Selection.End(xlToRight).Select
    Selection.End(xlToLeft).Select
    Range("M4").Select
    Selection.Copy
    Range("N4").Select
    Selection.End(xlDown).Select
    Range("M67").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("J67").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Range("I2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("H4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("H4").Select
    ActiveSheet.Paste
    Range("I4").Select
    Selection.End(xlDown).Select
    Range("H67").Select
    Selection.End(xlUp).Select
    Selection.Copy
    Range("I4").Select
    Selection.End(xlDown).Select
    Range("H67").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("B65").Select
    Selection.End(xlUp).Select
    Range("D1").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("C4").Select
    ActiveSheet.Paste
    Range("C4").Select
    Selection.Copy
    Range("D4").Select
    Selection.End(xlDown).Select
    Range("C67").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("C66").Select
    Selection.End(xlUp).Select
    Range("C1").Select
    Application.CutCopyMode = False
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "Target"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "Respuesta"
    Range("I1").Select
    Selection.ClearContents
    Range("I3").Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlUp).Select
    Range("CI1").Select
    Selection.End(xlToRight).Select
    Range("XFD4").Select
    Selection.End(xlToLeft).Select
    Range("CY4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CT68").Select
    ActiveSheet.Paste
    Range("CT67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CO68").Select
    ActiveSheet.Paste
    Range("CO67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CJ68").Select
    ActiveSheet.Paste
    Range("CJ67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CE68").Select
    ActiveSheet.Paste
    Range("CE67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("BZ68").Select
    ActiveSheet.Paste
    Range("BZ67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BU68").Select
    ActiveSheet.Paste
    Range("BU66").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BP68").Select
    ActiveSheet.Paste
    Range("BP67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BK68").Select
    ActiveSheet.Paste
    Range("BK67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BF68").Select
    ActiveSheet.Paste
    Range("BF67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BA68").Select
    ActiveSheet.Paste
    Range("BA67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("AV68").Select
    ActiveSheet.Paste
    Range("AV67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Selection.End(xlToLeft).Select
    Range("B4").Select
    Selection.End(xlToRight).Select
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AQ68").Select
    ActiveSheet.Paste
    Range("AQ67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AL68").Select
    ActiveSheet.Paste
    Range("AL67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AG68").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AB68").Select
    ActiveSheet.Paste
    Range("AB67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("W68").Select
    ActiveSheet.Paste
    Range("W67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("R68").Select
    ActiveSheet.Paste
    Range("R67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("M68").Select
    ActiveSheet.Paste
    Range("M67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("H68").Select
    ActiveSheet.Paste
    Range("H67").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("C68").Select
    ActiveSheet.Paste
    Range("C67").Select
    Selection.End(xlUp).Select
    Range("A61").Select
    Selection.End(xlUp).Select
    Range("A4").Select
    Selection.EntireColumn.Insert
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Variable"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "CONSUMO INDIVIDUO"
    Range("A4").Select
    Selection.Copy
    Range("B4").Select
    Application.CutCopyMode = False
    Range("A4").Select
    Selection.AutoFill Destination:=Range("A4:A1347")
    Range("A4:A1347").Select
    Range("B4:C4").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=21
    Selection.AutoFill Destination:=Range("B4:C1347"), Type:=xlFillDefault
    Range("B4:C1347").Select
    Range("C1343").Select
    Selection.End(xlUp).Select
    Range("I1:I6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlToLeft
    Range("H2").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.Delete Shift:=xlUp
    Range("B1").Select
    CLEARCONTROLER(0404)
    Else
    MsgBox ERROR
End If
End Sub








'==========================================| FIN CONSUMOINDIVIDUO | ========================================


'==========================================|  CONSUMOINDIVIDUO_MARCAS | ========================================

Sub CONSUMOINDIVIDUOMARCA()


Dim NOMBRE As String
Dim ERROR As String
Dim OK As String


ERROR = "LA MACRO SELECCIONADA NO ES VALIDA"
NOMBRE = "CONSUMO INDIVUO MARCAS"
OK = "OK"


If ActiveSheet.Name = NOMBRE Then


'
' BUILD Macro
'

'
    CONTROLADOR(0505)   
    Range("A4").Select
    Selection.EntireColumn.Insert
    Range("B4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A4").Select
    ActiveSheet.Paste
    Range("B1").Select
    Application.CutCopyMode = False
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("A4").Select
    Selection.Copy
    Range("A5:A11").Select
    ActiveSheet.Paste
    Range("A12").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A13:A34").Select
    ActiveSheet.Paste
    Range("A15").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("A35").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A36:A78").Select
    ActiveSheet.Paste
    Range("A79").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A80:A82").Select
    ActiveSheet.Paste
    Range("A83").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A84:A106").Select
    ActiveSheet.Paste
    Range("A91").Select
    Application.CutCopyMode = False
    Range("A107").Select
    Selection.Copy
    Range("A108:A125").Select
    ActiveSheet.Paste
    Range("A126").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A126:A128").Select
    ActiveSheet.Paste
    Range("A129").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A130").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("A129").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A130:A159").Select
    Range("A130:A158").Select
    ActiveWindow.SmallScroll Down:=15
    ActiveSheet.Paste
    Range("A159").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A160:A203").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=45
    Range("A204").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A204:A225").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=21
    Range("A226").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A227:A268").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=39
    Range("A270").Select
    Application.CutCopyMode = False
    Range("A269").Select
    Selection.Copy
    Range("A270:A300").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=27
    Range("A301").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A301:A324").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=24
    Range("A325").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A325:A349").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=27
    Range("A350").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A350:A367").Select
    ActiveSheet.Paste
    Range("A368").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A368:A414").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=48
    Range("A415").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A415:A424").Select
    ActiveSheet.Paste
    Range("A425").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A425:A435").Select
    ActiveSheet.Paste
    Range("A436").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A436:A443").Select
    ActiveSheet.Paste
    Range("A444").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A445:A460").Select
    ActiveSheet.Paste
    Range("A461").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A461:A468").Select
    ActiveSheet.Paste
    Range("A469").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A470:A475").Select
    ActiveSheet.Paste
    Range("A476").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A477:A505").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=33
    Range("A506").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A506:A527").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=24
    Range("A528").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A528:A566").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=42
    Range("A567").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A567:A605").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=45
    Range("A606").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A607:A637").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=33
    Range("A638").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A639:A664").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=24
    Range("A665").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A665:A668").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=12
    Range("A669").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A669:A713").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=48
    Range("A714").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A714:A726").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=9
    Range("A727").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A727:A785").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=57
    Range("A786").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A786:A850").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=69
    Range("A851").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A851:A909").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=60
    Range("A910").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A910:A916").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=6
    Range("A917").Select
    Application.CutCopyMode = False
    Range("A910").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A910:A927").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=18
    Range("A928").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A928:A969").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=45
    Range("A970").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A970:A1016").Select
    ActiveWindow.SmallScroll Down:=-9
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=54
    Range("A1017").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1017:A1029").Select
    ActiveSheet.Paste
    Range("A1029").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1030").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1030:A1070").Select
    ActiveSheet.Paste
    Range("A1071").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1071:A1085").Select
    ActiveSheet.Paste
    Range("A1086").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1086:A1106").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=27
    Range("A1107").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.End(xlDown).Select
    ActiveWindow.SmallScroll Down:=-465
    Range("A1107").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1107:A1142").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=39
    Range("A1143").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1143:A1169").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=30
    Range("A1170").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1170:A1184").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=12
    Range("A1185").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1185:A1217").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=39
    Range("A1218").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1218:A1241").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=24
    Range("A1242").Select
    Application.CutCopyMode = False
    Range("A1218").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1218:A1240").Select
    Range("A1240").Activate
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=12
    Range("A1241").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1241:A1243").Select
    ActiveSheet.Paste
    Range("A1244").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1245:A1254").Select
    ActiveSheet.Paste
    Range("A1255").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1256:A1300").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=45
    Range("A1301").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1301:A1306").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=12
    Range("A1307").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1307:A1315").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=9
    Range("A1316").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1316:A1334").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=21
    Range("A1335").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1335:A1342").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=12
    Range("A1343").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1343:A1350").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=9
    Range("A1351").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1351:A1361").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=6
    Range("A1362").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1362:A1375").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=15
    Range("A1376").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1376:A1385").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=12
    Range("A1386").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1386:A1404").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=12
    Range("A1405").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1405:A1423").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=15
    Range("A1424").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1424:A1453").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=33
    Range("A1454").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1454:A1462").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=9
    Range("A1463").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1463:A1476").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=12
    Range("A1477").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1477:A1483").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=9
    Range("A1484").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1484:A1487").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=9
    Range("A1488").Select
    ActiveWindow.SmallScroll Down:=-6
    Range("A1484").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1484:A1530").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=51
    Range("A1531").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1531:A1556").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=30
    Range("A1557").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1557:A1568").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=15
    Range("A1569").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1569:A1575").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=9
    Range("A1576").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1576:A1589").Select
    ActiveSheet.Paste
    Range("C1581").Select
    Application.CutCopyMode = False
    Range("B1583").Select
    ActiveWindow.SmallScroll Down:=-24
    Range("B1560").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("C2:F2").Select
    Selection.Cut
    Range("C4").Select
    Range("C4").Select
    Range("A4").Select
    Application.CutCopyMode = False
    Range("C1").Select
    Selection.AutoFilter
    Range("C4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveWindow.SmallScroll Down:=-48
    ActiveWindow.ScrollRow = 1427
    ActiveWindow.ScrollRow = 1347
    ActiveWindow.ScrollRow = 1046
    ActiveWindow.ScrollRow = 980
    ActiveWindow.ScrollRow = 839
    ActiveWindow.ScrollRow = 799
    ActiveWindow.ScrollRow = 669
    ActiveWindow.ScrollRow = 652
    ActiveWindow.ScrollRow = 592
    ActiveWindow.ScrollRow = 582
    ActiveWindow.ScrollRow = 562
    ActiveWindow.ScrollRow = 549
    ActiveWindow.ScrollRow = 529
    ActiveWindow.ScrollRow = 508
    ActiveWindow.ScrollRow = 485
    ActiveWindow.ScrollRow = 442
    ActiveWindow.ScrollRow = 422
    ActiveWindow.ScrollRow = 365
    ActiveWindow.ScrollRow = 341
    ActiveWindow.ScrollRow = 288
    ActiveWindow.ScrollRow = 258
    ActiveWindow.ScrollRow = 148
    ActiveWindow.ScrollRow = 124
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 1
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "filtr"
    Range("C4").Select
    Selection.AutoFilter
    Range("C3").Select
    Selection.AutoFilter
    Selection.ClearContents
    ActiveSheet.Range("$A$1:$CH$1589").AutoFilter Field:=3, Criteria1:="="
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.AutoFilter.ApplyFilter
    Selection.AutoFilter
    Range("C2").Select
    Selection.Cut
    Selection.End(xlToRight).Select
    Range("C2:F2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("C3").Select
    ActiveSheet.Paste
    Range("C4").Select
    ActiveWindow.SmallScroll Down:=30
    Range("C38").Select
    Selection.End(xlDown).Select
    Range("C1519").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("B1").Select
    Selection.Cut
    Range("A3").Select
    ActiveSheet.Paste
    Range("C3").Select
    Selection.EntireColumn.Insert
    Range("D1").Select
    Selection.Cut
    Range("C4").Select
    ActiveSheet.Paste
    Range("H4").Select
    Selection.EntireColumn.Insert
    Range("I2").Select
    Selection.Cut
    Range("H4").Select
    ActiveSheet.Paste
    Range("M4").Select
    Selection.EntireColumn.Insert
    Range("N2").Select
    Selection.Cut
    Range("M4").Select
    ActiveSheet.Paste
    Range("R2").Select
    Selection.EntireColumn.Insert
    Range("S2").Select
    Selection.Cut
    Range("R4").Select
    ActiveSheet.Paste
    Range("G4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("L4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("V4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("W4").Select
    Selection.EntireColumn.Insert
    Range("AB4").Select
    Selection.EntireColumn.Insert
    Range("AG4").Select
    Selection.EntireColumn.Insert
    Range("AL4").Select
    Selection.EntireColumn.Insert
    Range("AQ4").Select
    Selection.EntireColumn.Insert
    Range("AV4").Select
    Selection.EntireColumn.Insert
    Range("BA3").Select
    Selection.EntireColumn.Insert
    Range("BF3").Select
    Selection.EntireColumn.Insert
    Range("BK3").Select
    Selection.EntireColumn.Insert
    Range("BP3").Select
    Selection.EntireColumn.Insert
    Range("BU3").Select
    Selection.EntireColumn.Insert
    Range("BZ3").Select
    Selection.EntireColumn.Insert
    Range("CE3").Select
    Selection.EntireColumn.Insert
    Range("CJ3").Select
    Selection.EntireColumn.Insert
    Range("CO3").Select
    Selection.EntireColumn.Insert
    Range("CT3").Select
    Selection.EntireColumn.Insert
    Range("CY3").Select
    Selection.EntireColumn.Insert
    Range("CZ2").Select
    Selection.Cut
    Range("CY4").Select
    ActiveSheet.Paste
    Range("CY4").Select
    Selection.AutoFill Destination:=Range("CY4:CY1518")
    Range("CY4:CY1518").Select
    Range("DC4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("DB4").Select
    Selection.End(xlToLeft).Select
    Range("CX4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("CS4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("CU2").Select
    Selection.Cut
    Range("CT4").Select
    ActiveSheet.Paste
    Range("CT4").Select
    Selection.AutoFill Destination:=Range("CT4:CT1518")
    Range("CT4:CT1518").Select
    Range("CS4").Select
    Selection.End(xlToLeft).Select
    Range("CN4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("CP2").Select
    Selection.Cut
    Range("CO4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("CO4:CO1518")
    Range("CO4:CO1518").Select
    Range("CK2").Select
    Selection.Cut
    Range("CJ4").Select
    ActiveSheet.Paste
    Range("CJ4").Select
    Selection.AutoFill Destination:=Range("CJ4:CJ1518")
    Range("CJ4:CJ1518").Select
    Range("CI4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("CF2").Select
    Selection.Cut
    Range("CE4").Select
    ActiveSheet.Paste
    Range("CD4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("CE4").Select
    Selection.AutoFill Destination:=Range("CE4:CE1518")
    Range("CE4:CE1518").Select
    Range("CA2").Select
    Selection.Copy
    Range("BZ4").Select
    ActiveSheet.Paste
    Range("BY4").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "0"
    Range("BZ4").Select
    Selection.AutoFill Destination:=Range("BZ4:BZ1518")
    Range("BZ4:BZ1518").Select
    Range("BV2").Select
    Selection.Copy
    Range("BU4").Select
    ActiveSheet.Paste
    Range("BT4").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "0"
    Range("BV2").Select
    Selection.ClearContents
    Range("BU4").Select
    Selection.AutoFill Destination:=Range("BU4:BU1518")
    Range("BU4:BU1518").Select
    Range("CA2").Select
    Selection.ClearContents
    Range("BQ2").Select
    Selection.Cut
    Range("BP4").Select
    ActiveSheet.Paste
    Range("BO4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("BP4").Select
    Selection.AutoFill Destination:=Range("BP4:BP1518")
    Range("BP4:BP1518").Select
    Range("BL2").Select
    Selection.Copy
    Range("BK4").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("BK4:BK1518")
    Range("BK4:BK1518").Select
    Range("BL2").Select
    Selection.ClearContents
    Range("BJ4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("BG2").Select
    Selection.Cut
    Range("BF4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("BG4").Select
    Selection.End(xlDown).Select
    Range("BF1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BF1517").Select
    Selection.End(xlUp).Select
    Range("BB2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BA4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("BA4:BA1518")
    Range("BA4:BA1518").Select
    Range("AZ4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("AW1").Select
    Selection.Cut
    Range("AW2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("AW4").Select
    Selection.End(xlDown).Select
    Range("AV1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AQ1518").Select
    Selection.End(xlUp).Select
    Range("AR2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AQ4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("AR4").Select
    Selection.End(xlDown).Select
    Range("AQ1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AM1518").Select
    Selection.End(xlUp).Select
    Range("AM2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AL4").Select
    ActiveSheet.Paste
    Range("AL4").Select
    Selection.Copy
    Range("AM4").Select
    Selection.End(xlDown).Select
    Range("AL1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AH1518").Select
    Selection.End(xlUp).Select
    Range("AH2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AG4").Select
    ActiveSheet.Paste
    Range("AG4").Select
    Selection.Copy
    Range("AH4").Select
    Selection.End(xlDown).Select
    Range("AG1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AC1517").Select
    Selection.End(xlUp).Select
    Range("AC2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AB4").Select
    ActiveSheet.Paste
    Range("AA4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("AB4").Select
    Selection.Copy
    Range("AC4").Select
    Selection.End(xlDown).Select
    Range("AB1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("W1518").Select
    Selection.End(xlUp).Select
    Range("X2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("W4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("W4:W1518")
    Range("W4:W1518").Select
    Range("R4").Select
    Selection.Copy
    Range("S4").Select
    Selection.End(xlDown).Select
    Range("R1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("M1517").Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("M4:M1518")
    Range("M4:M1518").Select
    Range("H4").Select
    Selection.AutoFill Destination:=Range("H4:H1518")
    Range("H4:H1518").Select
    Range("C4").Select
    Selection.AutoFill Destination:=Range("C4:C1518")
    Range("C4:C1518").Select
    Range("A4").Select
    Selection.EntireColumn.Insert
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Variable"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "CONSUMO INDIVIDUO MARCAS"
    Range("A4").Select
    Selection.Copy
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("A1518").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("B1").Select
    Application.CutCopyMode = False
    Range("D3").Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlDown).Select
    Selection.End(xlToRight).Select
    Range("DD1516").Select
    Selection.End(xlUp).Select
    Range("DC1509").Select
    Selection.End(xlUp).Select
    Range("CZ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CU4").Select
    Selection.End(xlDown).Select
    Range("CU1519").Select
    ActiveSheet.Paste
    Range("CU1518").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CP4").Select
    Selection.End(xlDown).Select
    Range("CP1519").Select
    ActiveSheet.Paste
    Range("CP1518").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CK4").Select
    Selection.End(xlDown).Select
    Range("CK1519").Select
    ActiveSheet.Paste
    Range("CK1518").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CF4").Select
    Selection.End(xlDown).Select
    Range("CF1519").Select
    ActiveSheet.Paste
    Range("CF1518").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CA4").Select
    Selection.End(xlDown).Select
    Range("CA1519").Select
    ActiveSheet.Paste
    Range("CA1518").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Selection.End(xlToLeft).Select
    Range("BF4").Select
    Selection.End(xlToRight).Select
    Range("BH4").Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Range("XFC4").Select
    Selection.End(xlToLeft).Select
    Range("BV4").Select
    Selection.End(xlDown).Select
    Range("BV1519").Select
    ActiveSheet.Paste
    Range("BV1518").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BQ4").Select
    Selection.End(xlDown).Select
    Range("BQ1519").Select
    ActiveSheet.Paste
    Range("BQ1518").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BL4").Select
    Selection.End(xlDown).Select
    Range("BL1519").Select
    ActiveSheet.Paste
    Range("BL1518").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BG4").Select
    Selection.End(xlDown).Select
    Range("BG1519").Select
    ActiveSheet.Paste
    Range("BG1518").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Selection.End(xlToLeft).Select
    Range("BB4").Select
    Selection.End(xlDown).Select
    Range("BB1519").Select
    ActiveSheet.Paste
    Range("BB1518").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BB4").Select
    Application.CutCopyMode = False
    Range(Selection, Selection.End(xlToRight)).Select
    Range("BF4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("BB4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AW4").Select
    Selection.End(xlDown).Select
    Range("AW1519").Select
    ActiveSheet.Paste
    Range("AW1518").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AR4").Select
    Selection.End(xlDown).Select
    Range("AR1519").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("AV4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("AR4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AM4").Select
    Selection.End(xlDown).Select
    Range("AM1519").Select
    ActiveSheet.Paste
    Range("AM1518").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("AM3").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("AO4").Select
    Selection.End(xlToRight).Select
    Range("AQ4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("AM4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AI4").Select
    Selection.End(xlDown).Select
    Range("AI1517").Select
    Selection.End(xlUp).Select
    Range("AH3").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("AH1519").Select
    ActiveSheet.Paste
    
     Range("AH1517").Select
    Selection.End(xlUp).Select
    Range("AH4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("AH4:AL4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AC4").Select
    Selection.End(xlDown).Select
    Range("AC1519").Select
    ActiveSheet.Paste
    Range("AC1518").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("AC4:AG4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("X4").Select
    Selection.End(xlDown).Select
    Range("X1519").Select
    ActiveSheet.Paste
    Range("X1518").Select
    Selection.End(xlUp).Select
    Range("X4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("S4").Select
    Selection.End(xlDown).Select
    Range("S1519").Select
    ActiveSheet.Paste
    Range("S1518").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("N4").Select
    Selection.End(xlDown).Select
    Range("N1519").Select
    ActiveSheet.Paste
    Range("N1518").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range("N4:R4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("I5").Select
    Selection.End(xlDown).Select
    Range("I1519").Select
    ActiveSheet.Paste
    Range("I1518").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("D4").Select
    Selection.End(xlDown).Select
    Range("D1519").Select
    ActiveSheet.Paste
    Range("D1518").Select
    Selection.End(xlUp).Select
    Range("I1:I7").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlToLeft
    Range("H2").Select
    Selection.End(xlToLeft).Select
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Rows("1:2").Select
    Range("A2").Activate
    Selection.Delete Shift:=xlUp
    Range("A2").Select
    Selection.End(xlDown).Select
    Range("A1514").Select
    Selection.End(xlUp).Select
    Range("A2:C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A3").Select
    Selection.End(xlDown).Select
    Range("A1517").Select
    ActiveSheet.Paste
    Range("A1518").Select
    Selection.End(xlDown).Select
    Range("A3032").Select
    ActiveSheet.Paste
    Range("A3033").Select
    Selection.End(xlDown).Select
    Range("A4547").Select
    ActiveSheet.Paste
    Range("A4548").Select
    Selection.End(xlDown).Select
    Range("A6062").Select
    ActiveSheet.Paste
    Range("A6063").Select
    Selection.End(xlDown).Select
    Range("A7577").Select
    ActiveSheet.Paste
    Range("A7578").Select
    Selection.End(xlDown).Select
    Range("A9092").Select
    ActiveSheet.Paste
    Range("A9093").Select
    Selection.End(xlDown).Select
    Range("A10607").Select
    ActiveSheet.Paste
    Range("A10608").Select
    Selection.End(xlDown).Select
    Range("A12122").Select
    ActiveSheet.Paste
    Range("A12123").Select
    Selection.End(xlDown).Select
    Range("A13637").Select
    ActiveSheet.Paste
    Range("A13638").Select
    Selection.End(xlDown).Select
    Range("A15152").Select
    ActiveSheet.Paste
    Range("A15153").Select
    Selection.End(xlDown).Select
    Range("A16667").Select
    ActiveSheet.Paste
    Range("A16668").Select
    Selection.End(xlDown).Select
    Range("A18182").Select
    ActiveSheet.Paste
    Range("A18183").Select
    Selection.End(xlDown).Select
    Range("A19697").Select
    ActiveSheet.Paste
    Range("A19698").Select
    Selection.End(xlDown).Select
    Range("A21212").Select
    ActiveSheet.Paste
    Range("A21213").Select
    Selection.End(xlDown).Select
    Range("A22727").Select
    ActiveSheet.Paste
    Range("A22728").Select
    Selection.End(xlDown).Select
    Range("A24242").Select
    ActiveSheet.Paste
    Range("A24243").Select
    Selection.End(xlDown).Select
    Range("A25757").Select
    ActiveSheet.Paste
    Range("A25758").Select
    Selection.End(xlDown).Select
    Range("A27272").Select
    ActiveSheet.Paste
    Range("A27273").Select
    Selection.End(xlDown).Select
    Range("A28787").Select
    ActiveSheet.Paste
    Range("A28788").Select
    Selection.End(xlDown).Select
    Range("A30302").Select
    ActiveSheet.Paste
    Range("A30303").Select
    Selection.End(xlDown).Select
    Range("D31805").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("C1").Select
    Application.CutCopyMode = False
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Respuesta"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Target"
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace What:="", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("G31803").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select

    Range("M9").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("L1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlToLeft
    Range("I1:K1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlToLeft
    Range("I2").Select
    Selection.End(xlDown).Select
    Range("H1048576").Select
    Selection.End(xlUp).Select
    Range("E31815").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    
    Range("I4").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlToLeft
    Range("H4").Select

 Range("I2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("I1:XFD2").Select
    Range("I2").Activate
    Selection.Delete Shift:=xlToLeft
    Range("H2").Select
    Selection.End(xlToLeft).Select
    CLEARCONTROLER(0505)
Else
    MsgBox ERROR
End If
'FIN MACRO CONSUMO INDIVIDUO

   
End Sub


'==========================================| FIN CONSUMOINDIVIDUO_MARCAS | ========================================


'==============================================| CONSUMO HOGAR | =====================================================================

Sub CONSUMOHOGAR()
Attribute CONSUMOHOGARREP.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CONSUMOHOGARREP Macro
'




Dim NOMBRE As String
Dim ERROR As String
Dim OK As String


ERROR = "LA MACRO SELECCIONADA NO ES VALIDA"
NOMBRE = "CONSUMO HOGAR"
OK = "OK"



If ActiveSheet.Name = NOMBRE Then

    CONTROLADOR(0606)
    Range("A4").Select
    Selection.EntireColumn.Insert
    Range("B4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A4").Select
    ActiveSheet.Paste
    Range("B5").Select
    Application.CutCopyMode = False
    Range("C5").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A25").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A5:A25").Select
    Range("A25").Activate
    Range("A6:A25").Select
    Range("A25").Activate
    Selection.ClearContents
    Range("C25").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A31").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A27:A31").Select
    Range("A31").Activate
    Selection.ClearContents
    Range("C31").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A40").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A33:A40").Select
    Range("A40").Activate
    Selection.ClearContents
    Range("C40").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A46").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A42:A46").Select
    Range("A46").Activate
    Selection.ClearContents
    Range("C46").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A57").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A48:A57").Select
    Range("A57").Activate
    Selection.ClearContents
    Range("C57").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A74").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A59:A74").Select
    Range("A74").Activate
    Selection.ClearContents
    Range("C74").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A76").Select
    Selection.ClearContents
    Range("C77").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A81").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A81").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A78:A81").Select
    Range("A81").Activate
    Selection.ClearContents
    Range("C81").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A89").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A83:A89").Select
    Range("A89").Activate
    Selection.ClearContents
    Range("C89").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A96").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A91:A96").Select
    Range("A96").Activate
    Selection.ClearContents
    Range("C96").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A103").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A103").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A98:A103").Select
    Range("A103").Activate
    Selection.ClearContents
    Range("B102").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A5").Select
    Selection.AutoFill Destination:=Range("A5:A25")
    Range("A5:A25").Select
    ActiveWindow.SmallScroll Down:=18
    Range("A26").Select
    Selection.AutoFill Destination:=Range("A26:A31")
    Range("A26:A31").Select
    ActiveWindow.SmallScroll Down:=9
    Range("A32").Select
    Selection.AutoFill Destination:=Range("A32:A40")
    Range("A32:A40").Select
    ActiveWindow.SmallScroll Down:=9
    Range("A41").Select
    Selection.AutoFill Destination:=Range("A41:A46")
    Range("A41:A46").Select
    ActiveWindow.SmallScroll Down:=6
    Range("A47").Select
    Selection.AutoFill Destination:=Range("A47:A57")
    Range("A47:A57").Select
    ActiveWindow.SmallScroll Down:=9
    Range("A58").Select
    Selection.AutoFill Destination:=Range("A58:A74")
    Range("A58:A74").Select
    ActiveWindow.SmallScroll Down:=12
    Range("A75").Select
    Selection.AutoFill Destination:=Range("A75:A76")
    Range("A75:A76").Select
    ActiveWindow.SmallScroll Down:=9
    Range("A77").Select
    Selection.AutoFill Destination:=Range("A77:A81")
    Range("A77:A81").Select
    Range("A82").Select
    Selection.AutoFill Destination:=Range("A82:A89")
    Range("A82:A89").Select
    ActiveWindow.SmallScroll Down:=9
    Range("A90").Select
    Selection.AutoFill Destination:=Range("A90:A96")
    Range("A90:A96").Select
    Range("A97").Select
    Selection.AutoFill Destination:=Range("A97:A103")
    Range("A97:A103").Select
    Range("B91").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("C4").Select
    Selection.EntireColumn.Insert
    Range("H4").Select
    Selection.EntireColumn.Insert
    Range("M4").Select
    Selection.EntireColumn.Insert
    Range("R4").Select
    Selection.EntireColumn.Insert
    Range("W4").Select
    Selection.EntireColumn.Insert
    Range("AB4").Select
    Selection.EntireColumn.Insert
    Range("AG4").Select
    Selection.EntireColumn.Insert
    Range("AL4").Select
    Selection.EntireColumn.Insert
    Range("AQ4").Select
    Selection.EntireColumn.Insert
    Range("AV4").Select
    Selection.EntireColumn.Insert
    Range("BA4").Select
    Selection.EntireColumn.Insert
    Range("BF4").Select
    Selection.EntireColumn.Insert
    Range("BK4").Select
    Selection.EntireColumn.Insert
    Range("BP4").Select
    Selection.EntireColumn.Insert
    Range("BU4").Select
    Selection.EntireColumn.Insert
    Range("BZ4").Select
    Selection.EntireColumn.Insert
    Range("CE4").Select
    Selection.EntireColumn.Insert
    Range("CJ4").Select
    Selection.EntireColumn.Insert
    Range("CO4").Select
    Selection.EntireColumn.Insert
    Range("CT4").Select
    Selection.EntireColumn.Insert
    Range("CY4").Select
    Selection.EntireColumn.Insert
    Range("CV5").Select
    Selection.End(xlToLeft).Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Rows("4:4").Select
    Selection.Delete Shift:=xlUp
    Rows("4:4").Select
    Selection.Delete Shift:=xlUp
    Range("D4").Select
    Selection.End(xlDown).Select
    ActiveWindow.SmallScroll Down:=9
    Rows("24:24").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=6
    Rows("29:29").Select
    Selection.Delete Shift:=xlUp
    Range("D29").Select
    Selection.End(xlDown).Select
    ActiveWindow.SmallScroll Down:=9
    Rows("37:37").Select
    Selection.Delete Shift:=xlUp
    Range("D37").Select
    Selection.End(xlDown).Select
    ActiveWindow.SmallScroll Down:=3
    Rows("42:42").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=9
    Rows("52:52").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=21
    Rows("70:70").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=6
    Rows("74:74").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=6
    Rows("81:81").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=3
    Rows("87:87").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=-6
    Range("D91").Select
    Selection.End(xlUp).Select
    Rows("68:68").Select
    Selection.Delete Shift:=xlUp
    Range("D69").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    ActiveWindow.SmallScroll Down:=-30
    Range("I2").Select
    Selection.Cut
    Range("H4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("H4:H91")
    Range("H4:H91").Select
    Range("N2").Select
    Selection.Cut
    Range("M4").Select
    ActiveSheet.Paste
    Range("M4").Select
    ActiveCell.FormulaR1C1 = "HM 18 A 34 MB"
    Range("M4").Select
    Selection.AutoFill Destination:=Range("M4:M91")
    Range("M4:M91").Select
    Range("S2").Select
    Selection.Cut
    Range("R4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("S4").Select
    Selection.End(xlDown).Select
    Range("R91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("R90").Select
    Selection.End(xlUp).Select
    Range("X2").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("W4").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("W4:W91")
    Range("W4:W91").Select
    Range("X2").Select
    Selection.ClearContents
    Range("AC2").Select
    Selection.Cut
    Range("AB4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("AC4").Select
    Selection.End(xlDown).Select
    Range("AB91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AB90").Select
    Selection.End(xlUp).Select
    Range("AH2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AH4").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    ActiveSheet.Paste
    Range("AG4").Select
    Selection.Copy
    Range("AH4").Select
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range("AG91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AG90").Select
    Selection.End(xlUp).Select
    Range("AM2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AL4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("AM4").Select
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range("AL91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AQ91").Select
    Selection.End(xlUp).Select
    Range("AR2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AQ4").Select
    ActiveSheet.Paste
    Range("AQ4").Select
    Selection.Copy
    Range("AR4").Select
    Selection.End(xlDown).Select
    Range("AQ91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AQ90").Select
    Selection.End(xlUp).Select
    Range("AW4").Select
    Selection.End(xlUp).Select
    Range("AW2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    Range("AV4").Select
    Selection.Copy
    Range("AW4").Select
    Selection.End(xlDown).Select
    Range("AV91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AV90").Select
    Selection.End(xlUp).Select
    Range("BB2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BA4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("BA4:BA91")
    Range("BA4:BA91").Select
    Range("BG2").Select
    Selection.Cut
    Range("BF4").Select
    ActiveSheet.Paste
    Range("BF4").Select
    Selection.Copy
    Range("BG4").Select
    Selection.End(xlDown).Select
    Range("BF91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BF90").Select
    Selection.End(xlUp).Select
    Range("BL2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("BK4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("BK4:BK91")
    Range("BK4:BK91").Select
    Range("BQ2").Select
    Selection.Cut
    Range("BP4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("BP4:BP91")
    Range("BP4:BP91").Select
    Range("BV2").Select
    Selection.Cut
    Range("BU4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("BU4:BU91")
    Range("BU4:BU91").Select
    Range("CA2").Select
    Selection.Cut
    Range("BZ4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("BZ4:BZ91")
    Range("BZ4:BZ91").Select
    Range("CF2").Select
    Selection.Cut
    Range("CE4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("CE4:CE91")
    Range("CE4:CE91").Select
    Range("CK2").Select
    Selection.Cut
    Range("CJ4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("CJ4:CJ91")
    Range("CJ4:CJ91").Select
    Range("CP2").Select
    Selection.Cut
    Range("CO4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("CO4:CO91")
    Range("CO4:CO91").Select
    Range("CU2").Select
    Selection.Cut
    Range("CT4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("CT4:CT91")
    Range("CT4:CT91").Select
    Range("CZ2").Select
    Selection.Cut
    Range("CY4").Select
    ActiveSheet.Paste
    Selection.AutoFill Destination:=Range("CY4:CY91")
    Range("CY4:CY91").Select
    Range("DB3").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToRight).Select
    Range("CY4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CT92").Select
    ActiveSheet.Paste
    Range("CT91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CO92").Select
    ActiveSheet.Paste
    Range("CO91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CJ92").Select
    ActiveSheet.Paste
    Range("CJ91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CE92").Select
    ActiveSheet.Paste
    Range("CE91").Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("BZ92").Select
    ActiveSheet.Paste
    Range("BZ91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BU92").Select
    ActiveSheet.Paste
    Range("BU91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BP92").Select
    ActiveSheet.Paste
    Range("BP91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BK92").Select
    ActiveSheet.Paste
    Range("BK91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BF92").Select
    ActiveSheet.Paste
    Range("BF91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BA92").Select
    ActiveSheet.Paste
    Range("BA91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("AV92").Select
    ActiveSheet.Paste
    Range("AV91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AQ92").Select
    ActiveSheet.Paste
    Range("AQ91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("AQ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AL92").Select
    ActiveSheet.Paste
    Range("AL91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AG92").Select
    ActiveSheet.Paste
    Range("AG91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AB92").Select
    ActiveSheet.Paste
    Range("AB91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("W92").Select
    ActiveSheet.Paste
    Range("W91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("R92").Select
    ActiveSheet.Paste
    Range("R91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("M92").Select
    ActiveSheet.Paste
    Range("M91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("H92").Select
    ActiveSheet.Paste
    Range("H91").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("D4").Select
    Application.CutCopyMode = False
    Selection.EntireColumn.Insert
    Range("E1").Select
    Selection.Cut
    Range("D4").Select
    ActiveSheet.Paste
    Range("D4").Select
    Selection.AutoFill Destination:=Range("D4:D1763")
    Range("D4:D1763").Select
    Range("J4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("E4").Select
    Application.CutCopyMode = False
    Range("D4").Select
    Selection.Copy
    Range("E4").Select
    Selection.End(xlDown).Select
    Range("D91").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("H90").Select
    Selection.End(xlToRight).Select
    Range("I90").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("D4").Select
    Selection.End(xlDown).Select
    Range("D92").Select
    ActiveSheet.Paste
    Range("D91").Select
    Selection.End(xlUp).Select
    Range("J1").Select
    Selection.Cut
    Range("D2").Select
    ActiveSheet.Paste
    Range("B1").Select
    Selection.Cut
    Range("A3").Select
    ActiveSheet.Paste
    Range("D2:H2").Select
    Selection.Cut
    Range("D3").Select
    ActiveSheet.Paste
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "Respuesta"
    Range("A3").Select
    Selection.EntireColumn.Insert
    ActiveCell.FormulaR1C1 = "Variable"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "CONSUMO HOGAR"
    Range("A4").Select
    Selection.AutoFill Destination:=Range("A4:A91")
    Range("A4:A91").Select
    Range("B4:C4").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=12
    Range("D91").Select
    Selection.End(xlUp).Select
    Range("D3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("D10").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("D3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlToLeft
    Range("H2").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.Delete Shift:=xlUp
    Range("I1:I6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select

    ActiveWindow.SmallScroll Down:=12
    Range("B88").Select
    Selection.End(xlUp).Select
    Range("B2:C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=18
    Selection.AutoFill Destination:=Range("B2:C1849")
    Range("B2:C1849").Select
    Range("A89").Select
    Selection.AutoFill Destination:=Range("A89:A1849")
    Range("A89:A1849").Select
    ActiveWindow.SmallScroll Down:=27
    CLEARCONTROLER(0606)
    Else
    MsgBox ERROR
End If
End Sub



'===============================================| FIN CONSUMO HOGAR | ================================================================


'==============================================| CONSUMO HOGAR MARCAS | =====================================================================
Sub CONSUMOHOGARMARCAS()
'Attribute CHM_PREGUNTAS.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CHM_PREGUNTAS Macro
Dim NOMBRE As String
Dim ERROR As String
Dim OK As String


ERROR = "LA MACRO SELECCIONADA NO ES VALIDA"
NOMBRE = "CONSUMO HOGAR MARCAS"
OK = "OK"


If ActiveSheet.Name = NOMBRE Then
   

'

'
    CONTROLADOR(0707)   
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Range("A2").Select
    Selection.EntireColumn.Insert
    Range("B4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A4").Select
    ActiveSheet.Paste
    Range("A3").Select
    Application.CutCopyMode = False
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    Range("C4").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A21").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A5:A21").Select
    Range("A21").Activate
    Selection.ClearContents
    Range("C21").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A32").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A23:A32").Select
    Range("A32").Activate
    Selection.ClearContents
    Range("C32").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A49").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A34:A49").Select
    Range("A49").Activate
    Selection.ClearContents
    Range("C49").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A61").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A51:A61").Select
    Range("A61").Activate
    Selection.ClearContents
    Range("C62").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$CH$1554").AutoFilter Field:=3, Criteria1:="="
    ActiveWindow.SmallScroll Down:=-87
    ActiveSheet.Range("$A$1:$CH$1554").AutoFilter Field:=3, Criteria1:="<>"
    Range("A5").Select
    Selection.End(xlDown).Select
    Range("A63").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=63
    Range("C128").Select
    Selection.End(xlUp).Select
    Range("B1").Select
    Selection.AutoFilter
    Range("A4").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1375").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Range("B1541").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A4").Select
    Selection.Copy
    Selection.End(xlDown).Select
    Range("A21").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A22").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A23").Select
    Selection.End(xlDown).Select
    Range("A32").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A33").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A34").Select
    Selection.End(xlDown).Select
    Range("A49").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A50").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A51").Select
    Selection.End(xlDown).Select
    Range("A61").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A62").Select
    Selection.End(xlDown).Select
    Range("A76").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A63").Select
    Selection.End(xlDown).Select
    Range("A77").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A78").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A79").Select
    Selection.End(xlDown).Select
    Range("A93").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A94").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A95").Select
    Selection.End(xlDown).Select
    Range("A111").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A112").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A113").Select
    Selection.End(xlDown).Select
    Range("A122").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A123").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A124").Select
    Selection.End(xlDown).Select
    Range("A135").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A137").Select
    Selection.End(xlDown).Select
    Range("A142").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A143").Select
    Selection.End(xlDown).Select
    Range("A143").Select
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("A143:A154")
    Range("A143:A154").Select
    Range("A149").Select
    Selection.End(xlDown).Select
    Range("A156").Select
    ActiveWindow.SmallScroll Down:=15
    Range("A155").Select
    Selection.Copy
    Range("A156").Select
    Selection.End(xlDown).Select
    Range("A167").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A168").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A169").Select
    Selection.End(xlDown).Select
    Range("A185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A186").Select
    Selection.End(xlDown).Select
    Range("A193").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A187").Select
    Selection.End(xlDown).Select
    Range("A193").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A194").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A195").Select
    Selection.End(xlDown).Select
    Application.CutCopyMode = False
    ActiveWindow.SmallScroll Down:=-15

    Range("A194").Select
    Selection.AutoFill Destination:=Range("A194:A209")

    Range("A194:A209").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A210:A223")
    Range("A210:A223").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A224:A245")
    Range("A224:A245").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A246:A264")
    Range("A246:A264").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A265:A285")
    Range("A265:A285").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A286:A304")
    Range("A286:A304").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A305:A321")
    Range("A305:A321").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A322:A337")
    Range("A322:A337").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A338:A355")
    Range("A338:A355").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A356:A373")
    Range("A356:A373").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A374:A386")
    Range("A374:A386").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A387:A412")
    Range("A387:A412").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A413:A424")
    Range("A413:A424").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A425:A434")
    Range("A425:A431").Select
    Selection.End(xlDown).Select
    'Selection.AutoFill Destination:=Range("A432:A434")
    'Range("A432:A434").Select
    'Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A435:A447")
    Range("A435:A447").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A448:A459")
    Range("A448:A459").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A460:A477")
    Range("A460:A477").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A478:A489")
    Range("A478:A489").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A490:A503")
    Range("A490:A503").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A504:A525")
    Range("A504:A525").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A526:A546")
    Range("A526:A546").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A547:A576")
    Range("A547:A576").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A577:A603")
    Range("A577:A603").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A604:A624")
    Range("A604:A624").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A625:A641")
    Range("A625:A641").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A642:A662")
    Range("A642:A662").Select
    Selection.End(xlDown).Select
    'Selection.AutoFill Destination:=Range("A659:A662")
    'Range("A659:A662").Select
    'Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A663:A676")
    Range("A663:A676").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A677:A689")
    Range("A677:A689").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A690:A702")
    Range("A690:A702").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A703:A719")
    Range("A703:A719").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A720:A737")
    Range("A720:A737").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A738:A754")
    Range("A738:A754").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A755:A766")
    Range("A755:A766").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A767:A784")
    Range("A767:A784").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A785:A822")
    Range("A785:A822").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A823:A830")
    Range("A823:A830").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A831:A850")
    Range("A831:A850").Select
    Selection.End(xlDown).Select
    'Selection.AutoFill Destination:=Range("A847:A850")
    'Range("A847:A850").Select
    'Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A851:A866")
    Range("A851:A866").Select
    Selection.End(xlDown).Select
    'Selection.AutoFill Destination:=Range("A864:A866")
    'Range("A864:A866").Select
    'Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A867:A874")
    Range("A867:A874").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A875:A880")
    Range("A875:A880").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A881:A906")
    Range("A881:A906").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A907:A941")
    Range("A907:A941").Select
    Selection.End(xlDown).Select
    'Selection.AutoFill Destination:=Range("A939:A941")
    'Range("A939:A941").Select
    'Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A942:A984")
    Range("A942:A984").Select
    Selection.End(xlDown).Select

    Range("A986:A1014").Select
    Range("A1014").Activate
    Selection.ClearContents
    Range("A1015").Select

    ActiveWindow.SmallScroll Down:=-21
    Range("A985").Select

    Selection.AutoFill Destination:=Range("A985:A1014")
    Range("A985:A1014").Select
    Selection.End(xlDown).Select


    'Range("A1011").Select
    'Selection.AutoFill Destination:=Range("A1011:A1014")
    'Range("A1011:A1014").Select
    'Range("A1015").Select


    Selection.AutoFill Destination:=Range("A1015:A1024")
    Range("A1015:A1024").Select
    Selection.End(xlDown).Select
    'Range("A1025").Select


    Selection.AutoFill Destination:=Range("A1025:A1049")
    Range("A1025:A1049").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1050:A1068")
    Range("A1050:A1068").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1069:A1075")
    Range("A1069:A1075").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1076:A1101")
    Range("A1076:A1101").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1102:A1131")
    Range("A1102:A1131").Select
    Selection.End(xlDown).Select
    'Selection.AutoFill Destination:=Range("A1128:A1131")
    'Range("A1128:A1131").Select
    'Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1132:A1154")
    Range("A1132:A1154").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1155:A1186")
    Range("A1155:A1186").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1187:A1204")
    Range("A1187:A1204").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1205:A1213")
    Range("A1205:A1213").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1214:A1232")
    Range("A1214:A1232").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1233:A1247")
    Range("A1233:A1247").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1248:A1254")
    Range("A1248:A1254").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1255:A1280")
    Range("A1255:A1280").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1281:A1310")
    Range("A1281:A1310").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1311:A1325")
    Range("A1311:A1325").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1326:A1351")
    Range("A1326:A1351").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1352:A1365")
    Range("A1352:A1365").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1366:A1375")
    Range("A1366:A1375").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1376:A1388")
    Range("A1376:A1388").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1389:A1422")
    Range("A1389:A1422").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1423:A1433")
    Range("A1423:A1433").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1434:A1451")
    Range("A1434:A1451").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1452:A1477")
    Range("A1452:A1477").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1478:A1486")
    Range("A1478:A1486").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1487:A1505")
    Range("A1487:A1505").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1506:A1512")
    Range("A1506:A1512").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1513:A1531")
    Range("A1513:A1531").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1532:A1540")
    Range("A1532:A1540").Select
    Selection.End(xlDown).Select
    Selection.AutoFill Destination:=Range("A1541:A1554")
    Range("A1541:A1554").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A1048576").Select
    ActiveWindow.SmallScroll Down:=-42
    Selection.End(xlUp).Select
    ActiveWindow.SmallScroll Down:=-51
    Range("A1511").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A4").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$CH$1554").AutoFilter Field:=2, Criteria1:="="
    ActiveWindow.SmallScroll Down:=-18
    Selection.AutoFilter
    Range("C4").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$CH$1554").AutoFilter Field:=3, Criteria1:="="
    Range("A4:B4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Selection.AutoFilter
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("B1459").Select
    Selection.End(xlUp).Select
    Range("C2:F2").Select
    Selection.Cut
    Range("C3").Select
    ActiveSheet.Paste
    Range("C1").Select
    Selection.Cut
    Range("C3").Select
    Application.CutCopyMode = False
    Range("B1").Select
    Selection.Cut
    Range("A3").Select
    ActiveSheet.Paste
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "Respuesta"
    Range("B4").Select
    CSHMTARGET
    Else
    MsgBox ERROR
End If
End Sub

Attribute VB_Name = "M�dulo2"
Private Sub CSHMTARGET()
'Attribute UNIFICAR_TARGETS.VB_ProcData.VB_Invoke_Func = " \n14"
'
' UNIFICAR_TARGETS Macro
'

'
    Range("G3").Select
    Selection.EntireColumn.Insert
    Range("L2").Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToLeft).Select
    Range("CY2").Select
    Selection.Cut Destination:=Range("CX4")
    Range("CT2").Select
    Selection.Cut Destination:=Range("CS4")
    Range("CO2").Select
    Selection.Cut Destination:=Range("CN4")
    Range("CJ2").Select
    Selection.Cut Destination:=Range("CI4")
    Range("CE2").Select
    Selection.Cut Destination:=Range("CD4")
    Range("BZ2").Select
    Selection.Cut Destination:=Range("BY4")
    Range("BU2").Select
    Selection.Cut Destination:=Range("BT4")
    Range("BP2").Select
    Selection.Cut Destination:=Range("BO4")
    Range("BK2").Select
    Selection.Cut Destination:=Range("BJ4")
    Range("BF2").Select
    Selection.Cut Destination:=Range("BE4")
    Range("BA2").Select
    Selection.Cut Destination:=Range("AZ4")
    Range("AV2").Select
    Selection.Cut Destination:=Range("AU4")
    Range("AQ2").Select
    Selection.Cut Destination:=Range("AP4")
    Range("AL2").Select
    Selection.Cut Destination:=Range("AK4")
    Range("AG2").Select
    Selection.Cut Destination:=Range("AF4")
    Range("AB2").Select
    Selection.Cut Destination:=Range("AA4")
    Range("W2").Select
    Selection.Cut Destination:=Range("V4")
    Range("R2").Select
    Selection.Cut Destination:=Range("Q4")
    Range("M2").Select
    Selection.Cut Destination:=Range("L4")
    Range("H2").Select
    Selection.Cut Destination:=Range("G4")
    Range("C4").Select
    Selection.EntireColumn.Insert
    Range("D1").Select
    Range("D1").Select
    Selection.Cut Destination:=Range("C4")
    Range("C4").Select
    Selection.AutoFill Destination:=Range("C4:C1459")
    Range("C4:C1459").Select
    Range("H4").Select
    Selection.AutoFill Destination:=Range("H4:H1459")
    Range("H4:H1459").Select
    Range("M4").Select
    Selection.Copy
    Range("N4").Select
    Selection.End(xlDown).Select
    Range("M1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("M1458").Select
    Selection.End(xlUp).Select
    Range("R4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("S4").Select
    Selection.End(xlDown).Select
    Range("R1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("R1458").Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("X4").Select
    Selection.End(xlDown).Select
    Range("W1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("W1458").Select
    Selection.End(xlUp).Select
    Range("AB4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AC4").Select
    Selection.End(xlDown).Select
    Range("AB1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AB1458").Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AH4").Select
    Selection.End(xlDown).Select
    Range("AG1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AG1458").Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AM4").Select
    Selection.End(xlDown).Select
    Range("AL1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AL1458").Select
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AR4").Select
    Selection.End(xlDown).Select
    Range("AQ1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AQ1458").Select
    Selection.End(xlUp).Select
    Range("AV4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AW4").Select
    Selection.End(xlDown).Select
    Range("AV1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AV1458").Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BB4").Select
    Selection.End(xlDown).Select
    Range("BA1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BA1458").Select
    Selection.End(xlUp).Select
    Range("BG4").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToRight).Select
    Range("DC6").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range("CS5").Select
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Range("CP4").Select
    Selection.End(xlToLeft).Select
    Range("A6").Select
    Selection.End(xlToRight).Select
    Range("BF4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BG4").Select
    Selection.End(xlDown).Select
    Range("BF1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BF1458").Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BL4").Select
    Selection.End(xlDown).Select
    Range("BK1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BK1458").Select
    Selection.End(xlUp).Select
    Range("BO4").Select
    Selection.End(xlToRight).Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Range("CX4").Select
    Selection.End(xlToLeft).Select
    Range("B6").Select
    Selection.End(xlToRight).Select
    Range("BP4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BQ4").Select
    Selection.End(xlDown).Select
    Range("BP1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BP1458").Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BV4").Select
    Selection.End(xlDown).Select
    Range("BU1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BU1458").Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CA4").Select
    Selection.End(xlDown).Select
    Range("BZ1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BZ1458").Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CF4").Select
    Selection.End(xlDown).Select
    Range("CE1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CE1458").Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CK4").Select
    Selection.End(xlDown).Select
    Range("CJ1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CJ1458").Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CP4").Select
    Selection.End(xlDown).Select
    Range("CO1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CO1458").Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CU4").Select
    Selection.End(xlDown).Select
    Range("CT1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CT1458").Select
    Selection.End(xlUp).Select
    Range("CY4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CZ4").Select
    Selection.End(xlDown).Select
    Range("CY1459").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("CY1458").Select
    Selection.End(xlUp).Select
    Range("DE3").Select
    Selection.End(xlUp).Select
    Range("CV1").Select
    Application.CutCopyMode = False
    Range("CY4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CT1460").Select
    ActiveSheet.Paste
    Range("CT1459").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CO1460").Select
    ActiveSheet.Paste
    Range("CO1459").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CJ1460").Select
    ActiveSheet.Paste
    Range("CJ1459").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CE1460").Select
    ActiveSheet.Paste
    Range("CE1459").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("BZ1460").Select
    ActiveSheet.Paste
    Range("BZ1459").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BU1460").Select
    ActiveSheet.Paste
    Range("BU1459").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BP1460").Select
    ActiveSheet.Paste
    Range("BP1459").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BK1460").Select
    ActiveSheet.Paste
    Range("BK1459").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BF1460").Select
    ActiveSheet.Paste
    Range("BF1459").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BA1460").Select
    ActiveSheet.Paste
    Range("BA1459").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("AV1460").Select
    ActiveSheet.Paste
    Range("AV1459").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AQ1460").Select
    ActiveSheet.Paste
    Range("AQ1459").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AL1460").Select
    ActiveSheet.Paste
    Range("AL1459").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AG1460").Select
    ActiveSheet.Paste
    Range("AG1459").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AB1460").Select
    ActiveSheet.Paste
    Range("AB1459").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("W1460").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Cut
    Range("W4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("R1460").Select
    ActiveSheet.Paste
    Range("R1459").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("M1460").Select
    ActiveSheet.Paste
    Range("M1459").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("H1460").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("C1460").Select
    ActiveSheet.Paste
    Range("B1460").Select
    Selection.End(xlToLeft).Select
    Range("A1371").Select
    Selection.End(xlUp).Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("I1").Select
    Selection.Cut Destination:=Range("C3")
    Range("H1:H5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlToLeft
    Range("G2").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    Selection.EntireColumn.Insert
    ActiveCell.FormulaR1C1 = "Variable"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "CONSUMO HOGAR MARCAS"
    Range("A2").Select
    Selection.Copy
    Range("B2").Select
    Selection.End(xlDown).Select
    Range("A1457").Select
    Selection.End(xlUp).Select
    Range("B3").Select
    Selection.End(xlDown).Select
    Range("A1457").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A1").Select
    Application.CutCopyMode = False
    Range("A2:C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=18
    Selection.AutoFill Destination:=Range("A2:C30577"), Type:=xlFillDefault
    Range("A2:C30577").Select
    Range("B30566").Select
    Selection.End(xlUp).Select
    Range("B1").Select

    CLEARCONTROLER(0707)


End Sub






'==============================================| FIN CONSUMO HOGAR MARCAS | =====================================================================


'==============================================| SERVICIOS FINANCIEROS | =====================================================================
Sub SERVICIOSFINANCIEROS()
'Attribute SERVICIOSFINANCIEROSCORREGIDA.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SERVICIOSFINANCIEROSCORREGIDA Macro
'

'



Dim NOMBRE As String
Dim ERROR As String
Dim OK As String


ERROR = "LA MACRO SELECCIONADA NO ES VALIDA"
NOMBRE = "SERVICIOS FINANCIEROS"
OK = "OK"


If ActiveSheet.Name = NOMBRE Then
   

    CONTROLADOR(0808)
    Selection.EntireColumn.Insert
    Range("B4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A4").Select
    ActiveSheet.Paste
    Range("A3").Select
    Application.CutCopyMode = False
    Range("A6:A13").Select
    Selection.ClearContents
    Range("A15:A33").Select
    Selection.ClearContents
    Range("A16").Select
    Selection.End(xlDown).Select
    Range("A35:A52").Select
    Selection.ClearContents
    Range("C36").Select
    Selection.End(xlDown).Select
    Range("C53").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A77").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A54:A77").Select
    Range("A77").Activate
    Selection.ClearContents
    Range("C77").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A85").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A79:A85").Select
    Range("A85").Activate
    Selection.ClearContents
    Range("C85").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A107").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A87:A107").Select
    Range("A107").Activate
    Selection.ClearContents
    Range("C107").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A135").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A109:A135").Select
    Range("A135").Activate
    Selection.ClearContents
    Range("C135").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A160").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A137:A160").Select
    Range("A160").Activate
    Selection.ClearContents
    Range("C160").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range("A195").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A162:A195").Select
    Range("A195").Activate
    Selection.ClearContents
    Range("B195").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A1").Select
    Selection.End(xlUp).Select
    Range("A5").Select
    Selection.Copy
    Range("A6:A13").Select
    Range("A13").Activate
    ActiveSheet.Paste
    Range("A14").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C14").Select
    Selection.End(xlDown).Select
    Range("C15").Select
    Selection.End(xlDown).Select
    Range("A33").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A15:A33").Select
    Range("A33").Activate
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("A32").Select
    Selection.End(xlUp).Select
    Selection.Copy
    Range("C14").Select
    Selection.End(xlDown).Select
    Range("C15").Select
    Selection.End(xlDown).Select
    Range("A33").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A15:A33").Select
    Range("A33").Activate
    ActiveSheet.Paste
    Range("A34").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C34").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A52").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A35:A52").Select
    Range("A52").Activate
    ActiveSheet.Paste
    Range("B52").Select
    Selection.End(xlToRight).Select
    Range("CG52").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range("A53").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C53").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A77").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A54:A77").Select
    Range("A77").Activate
    ActiveSheet.Paste
    Range("C77").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A78").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C78").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A85").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A79:A85").Select
    Range("A85").Activate
    ActiveSheet.Paste
    Range("C85").Select
    Selection.End(xlDown).Select
    Range("A86").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C86").Select
    Selection.End(xlDown).Select
    Range("A87").Select
    Selection.End(xlDown).Select
    Range("A107").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A108").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A109").Select
    Selection.End(xlDown).Select
    Range("A135").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A136").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A137").Select
    Selection.End(xlDown).Select
    Range("A160").Select
    Range(Selection, Selection.End(xlUp)).Select
  
    Range("A160").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A159").Select
    Selection.End(xlUp).Select
    Selection.Copy
    Range("A137").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Range("A139").Select
    Selection.End(xlDown).Select
    Range("A161").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A162").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("B161").Select
    Selection.End(xlDown).Select
    Range("A195").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("A194").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Rows("5:5").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    Rows("13:13").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=15
    Rows("32:32").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=27
    Rows("50:50").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=18
    Rows("74:74").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=6
    Rows("81:81").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=18
    Rows("101:101").Select
    Rows("102:102").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=24
    Rows("129:129").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=24
    Rows("153:153").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=15
    Range("C168").Select
    Selection.End(xlUp).Select
    ActiveWindow.SmallScroll Down:=-21
    Rows("4:4").Select
    Selection.Delete Shift:=xlUp
    Range("C2:F2").Select
    Selection.Cut
    Range("C3").Select
    ActiveSheet.Paste
    Range("B1").Select
    Selection.Cut
    Range("A3").Select
    ActiveSheet.Paste
    Selection.End(xlToRight).Select
    Range("C3").Select
    Selection.EntireColumn.Insert
    Range("H3").Select
    Selection.EntireColumn.Insert
    Range("M3").Select
    Selection.EntireColumn.Insert
    Range("N2").Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Range("BA2").Select
    Selection.EntireColumn.Insert
    Range("BB2").Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Range("CE2").Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Range("CY2").Select
    Selection.EntireColumn.Insert
    Selection.End(xlToRight).Select
    Selection.End(xlToRight).Select
    Range("XFC2").Select
    Selection.End(xlToLeft).Select
    Range("CZ2").Select
    Selection.Cut
    Range("CY4").Select
    ActiveSheet.Paste
    Range("CU2").Select
    Selection.Cut
    Range("CT4").Select
    ActiveSheet.Paste
    Range("CP2").Select
    Selection.Cut
    Range("CO4").Select
    ActiveSheet.Paste
    Range("CM2").Select
    Selection.End(xlToLeft).Select
    Range("CK2").Select
    Selection.Cut
    Range("CJ4").Select
    ActiveSheet.Paste
    Range("CJ2").Select
    Selection.End(xlToLeft).Select
    Selection.Cut
    Range("CE4").Select
    ActiveSheet.Paste
    Range("CB2").Select
    Selection.End(xlToLeft).Select
    Selection.Cut
    Range("BZ4").Select
    ActiveSheet.Paste
    Range("BV2").Select
    Selection.Cut
    Range("BU4").Select
    ActiveSheet.Paste
    Range("BQ2").Select
    Selection.Cut
    Range("BP4").Select
    ActiveSheet.Paste
    Range("BL2").Select
    Selection.Cut
    Range("BK4").Select
    ActiveSheet.Paste
    Range("BG2").Select
    Selection.Cut
    Range("BF4").Select
    ActiveSheet.Paste
    Range("BB2").Select
    Selection.Cut
    Range("BA4").Select
    ActiveSheet.Paste
    Range("AW2").Select
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    Range("AR2").Select
    Selection.Cut
    Range("AQ4").Select
    ActiveSheet.Paste
    Range("AM2").Select
    Selection.Cut
    Range("AL4").Select
    ActiveSheet.Paste
    Range("AH2").Select
    Selection.Cut
    Range("AG4").Select
    ActiveSheet.Paste
    Range("AC2").Select
    Selection.Cut
    Range("AB4").Select
    ActiveSheet.Paste
    Range("X2").Select
    Selection.Cut
    Range("W4").Select
    ActiveSheet.Paste
    Range("S2").Select
    Selection.Cut
    Range("R4").Select
    ActiveSheet.Paste
    Range("N2").Select
    Selection.Cut
    Range("M4").Select
    ActiveSheet.Paste
    Range("I2").Select
    Selection.Cut
    Range("H4").Select
    ActiveSheet.Paste
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "Respuesta"
    Range("D1").Select
    Selection.Cut
    Range("C4").Select
    ActiveSheet.Paste
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "Target"
    Range("I1").Select
    Selection.ClearContents
    Range("C4").Select
    Selection.Copy
    Range("D5").Select
    Selection.End(xlDown).Select
    Range("C185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H4").Select
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("H4:H185")
    Range("H4:H185").Select
    Range("M4").Select
    Selection.AutoFill Destination:=Range("M4:M185")
    Range("M4:M185").Select
    Range("R4").Select
    Selection.Copy
    Range("S4").Select
    Selection.End(xlDown).Select
    Range("R185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("R184").Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("W4:W185")
    Range("W4:W185").Select
    Range("AB4").Select
    Selection.Copy
    Range("AC4").Select
    Selection.End(xlDown).Select
    Range("AB185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AB184").Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AG4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AH4").Select
    Selection.End(xlDown).Select
    Range("AG185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AG184").Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AM4").Select
    Selection.End(xlDown).Select
    Range("AL185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AL184").Select
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AR4").Select
    Selection.End(xlDown).Select
    Range("AQ185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AQ184").Select
    Selection.End(xlUp).Select
    Range("AV4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AW4").Select
    Selection.End(xlDown).Select
    Range("AV185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("AV184").Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("BA4:BA185")
    Range("BA4:BA185").Select
    Range("BF4").Select
    Selection.Copy
    Range("BG4").Select
    Selection.End(xlDown).Select
    Range("BF185").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BF184").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BG4").Select
    Selection.End(xlDown).Select
    Range("BF185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BF184").Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BL4").Select
    Selection.End(xlDown).Select
    Range("BK185").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Range("BK184").Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("BP4:BP185")
    Range("BP4:BP185").Select
    Range("BU4").Select
    Selection.AutoFill Destination:=Range("BU4:BU185")
    Range("BU4:BU185").Select
    Range("BZ4").Select
    Selection.AutoFill Destination:=Range("BZ4:BZ185")
    Range("BZ4:BZ185").Select
    Range("CJ4").Select
    Selection.AutoFill Destination:=Range("CJ4:CJ185")
    Range("CJ4:CJ185").Select
    Range("CE4").Select
    Selection.AutoFill Destination:=Range("CE4:CE185")
    Range("CE4:CE185").Select
    Range("CO4").Select
    Selection.AutoFill Destination:=Range("CO4:CO185")
    Range("CO4:CO185").Select
    Range("CT4").Select
    Selection.AutoFill Destination:=Range("CT4:CT185")
    Range("CT4:CT185").Select
    Range("CY4").Select
    Selection.AutoFill Destination:=Range("CY4:CY185")
    Range("CY4:CY185").Select
    Range("CY4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CT186").Select
    ActiveSheet.Paste
    Range("CT185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CO186").Select
    ActiveSheet.Paste
    Range("CO185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CJ186").Select
    ActiveSheet.Paste
    Range("CJ185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CE186").Select
    ActiveSheet.Paste
    Range("CE185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("BZ186").Select
    ActiveSheet.Paste
    Range("BZ186").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Application.CutCopyMode = False
    Selection.Copy
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BU186").Select
    Range("BZ186").Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Range("BZ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BU186").Select
    ActiveSheet.Paste
    Range("BU185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BP186").Select
    ActiveSheet.Paste
    Range("BP185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BK186").Select
    ActiveSheet.Paste
    Range("BK185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BF186").Select
    ActiveSheet.Paste
    Range("BF185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BA186").Select
    ActiveSheet.Paste
    Range("BA185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("AV186").Select
    ActiveSheet.Paste
    Range("AV185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AQ186").Select
    ActiveSheet.Paste
    Range("AQ185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AL186").Select
    ActiveSheet.Paste
    Range("AL185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AG186").Select
    ActiveSheet.Paste
    Range("AG185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AB186").Select
    ActiveSheet.Paste
    Range("AB185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("W186").Select
    ActiveSheet.Paste
    Range("W185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("R186").Select
    ActiveSheet.Paste
    Range("R185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("M186").Select
    ActiveSheet.Paste
    Range("M185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("H186").Select
    ActiveSheet.Paste
    Range("H185").Select
    Selection.End(xlUp).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("C186").Select
    ActiveSheet.Paste
    Range("C185").Select
    Selection.End(xlUp).Select
    Range("A4:B5").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=21
    Selection.AutoFill Destination:=Range("A4:B3825"), Type:=xlFillDefault
    Range("A4:B3825").Select
    Range("B3825").Select
    Selection.End(xlUp).Select
    Range("H3:DF3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlToLeft
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Range("A2").Select
    Selection.EntireColumn.Insert
    ActiveCell.FormulaR1C1 = "Variable"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "SERVICIOS FINANCIEROS"
    Range("A3").Select
    Selection.AutoFill Destination:=Range("A3:A3824")
    Range("A3:A3824").Select
    Range("A8").Select
    Selection.End(xlDown).Select
    ActiveWindow.SmallScroll Down:=6
    Selection.End(xlUp).Select

      Rows("1:1").Select
    Selection.Delete Shift:=xlUp

    CLEARCONTROLER(0808)

    Else
    MsgBox ERROR
End If
 
End Sub

'==============================================| FIN SERVICIOS FINANCIEROS | =====================================================================



'==============================================| TRANSPORTE FINANCIEROS | =====================================================================



Sub TRANSPORTE()

Dim NOMBRE As String
Dim ERROR As String
Dim OK As String


ERROR = "LA MACRO SELECCIONADA NO ES VALIDA"
NOMBRE = "TRANSPORTE"
OK = "OK"


If ActiveSheet.Name = NOMBRE Then




Attribute TRANSPORTE.VB_ProcData.VB_Invoke_Func = "j\n14"
'
' TRANSPORTE Macro
'
' Acceso directo: CTRL+j
'
    CONTROLADOR(0909)
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.ColumnWidth = 42.75
    Range("B5").Select
    Selection.Cut
    Range("A6").Select
    ActiveSheet.Paste
    Range("B18").Select
    Selection.Cut
    Range("A19").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=9
    Range("B21").Select
    Selection.Cut
    Range("A22").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=3
    Range("B35").Select
    Selection.Cut
    Range("A36").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=9
    Range("B38").Select
    Selection.Cut
    Range("A39").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=6
    Range("B46").Select
    Selection.Cut
    Range("A47").Select
    ActiveSheet.Paste
    Range("B50").Select
    Selection.Cut
    Range("A51").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=6
    Range("B54").Select
    Selection.Cut
    Range("A55").Select
    ActiveSheet.Paste
    Range("B58").Select
    Selection.Cut
    Range("A59").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=27
    ActiveWindow.ScrollRow = 60
    ActiveWindow.ScrollRow = 59
    ActiveWindow.ScrollRow = 58
    ActiveWindow.ScrollRow = 57
    ActiveWindow.ScrollRow = 56
    ActiveWindow.ScrollRow = 55
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 53
    ActiveWindow.ScrollRow = 52
    ActiveWindow.ScrollRow = 51
    ActiveWindow.ScrollRow = 49
    ActiveWindow.ScrollRow = 48
    ActiveWindow.ScrollRow = 47
    ActiveWindow.ScrollRow = 46
    ActiveWindow.ScrollRow = 45
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 42
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 38
    ActiveWindow.ScrollRow = 37
    ActiveWindow.ScrollRow = 36
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 34
    ActiveWindow.ScrollRow = 32
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 29
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 26
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    Rows("4:5").Select
    Selection.Delete Shift:=xlUp
    Rows("16:16").Select
    Selection.Delete Shift:=xlUp
    Rows("18:18").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=12
    Rows("31:31").Select
    Selection.Delete Shift:=xlUp
    Rows("33:33").Select
    Selection.Delete Shift:=xlUp
    Range("B35").Select
    ActiveWindow.SmallScroll Down:=6
    Rows("40:40").Select
    Selection.Delete Shift:=xlUp
    Rows("43:43").Select
    Selection.Delete Shift:=xlUp
    Rows("46:46").Select
    Selection.Delete Shift:=xlUp
    Range("A39").Select
    ActiveWindow.SmallScroll Down:=9
    Rows("49:49").Select
    Selection.Delete Shift:=xlUp
    Range("C52").Select
    ActiveWindow.SmallScroll Down:=75
    ActiveWindow.ScrollRow = 102
    ActiveWindow.ScrollRow = 101
    ActiveWindow.ScrollRow = 100
    ActiveWindow.ScrollRow = 99
    ActiveWindow.ScrollRow = 98
    ActiveWindow.ScrollRow = 97
    ActiveWindow.ScrollRow = 96
    ActiveWindow.ScrollRow = 94
    ActiveWindow.ScrollRow = 93
    ActiveWindow.ScrollRow = 91
    ActiveWindow.ScrollRow = 89
    ActiveWindow.ScrollRow = 87
    ActiveWindow.ScrollRow = 86
    ActiveWindow.ScrollRow = 84
    ActiveWindow.ScrollRow = 82
    ActiveWindow.ScrollRow = 78
    ActiveWindow.ScrollRow = 76
    ActiveWindow.ScrollRow = 74
    ActiveWindow.ScrollRow = 72
    ActiveWindow.ScrollRow = 71
    ActiveWindow.ScrollRow = 69
    ActiveWindow.ScrollRow = 67
    ActiveWindow.ScrollRow = 65
    ActiveWindow.ScrollRow = 63
    ActiveWindow.ScrollRow = 62
    ActiveWindow.ScrollRow = 60
    ActiveWindow.ScrollRow = 58
    ActiveWindow.ScrollRow = 55
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 52
    ActiveWindow.ScrollRow = 49
    ActiveWindow.ScrollRow = 45
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 36
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 1
    Range("A4").Select
    Selection.Copy
    Range("A5:A15").Select
    ActiveSheet.Paste
    Range("A16").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A17").Select
    ActiveSheet.Paste
    Range("A18").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A19").Select
    ActiveWindow.SmallScroll Down:=12
    Range("A19:A30").Select
    ActiveSheet.Paste
    Range("A31").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A32").Select
    ActiveSheet.Paste
    Range("A33").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A34:A39").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=9
    Range("A40").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A41:A42").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=6
    Range("A43").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A44:A45").Select
    ActiveSheet.Paste
    Range("A46").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A47:A48").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=9
    Range("A49").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A50").Select
    ActiveWindow.SmallScroll Down:=63
    Range("A50:A109").Select
    ActiveSheet.Paste
    Range("A49").Select
    Application.CutCopyMode = False
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 34
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 32
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 30
    ActiveWindow.ScrollRow = 29
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 25
    ActiveWindow.ScrollRow = 24
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 20
    ActiveWindow.ScrollRow = 19
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 16
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 14
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    Columns("A:A").ColumnWidth = 46.25
    Columns("B:B").ColumnWidth = 42
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("R:R").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("W:W").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    Columns("AB:AB").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AG:AG").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    Columns("AL:AL").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AQ:AQ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    Columns("AV:AV").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BA:BA").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    Columns("BF:BF").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BK:BK").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 59
    Columns("BP:BP").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BU:BU").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BZ:BZ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    Columns("CE:CE").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CJ:CJ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    Columns("CO:CO").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CT:CT").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CY:CY").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 90
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 1
    Columns("C:C").ColumnWidth = 26.25
    Range("D1").Select
    Selection.Cut
    Range("C4").Select
    ActiveSheet.Paste
    Range("I2").Select
    Selection.Cut
    Range("H4").Select
    ActiveSheet.Paste
    Range("N2").Select
    Selection.Cut
    Range("M4").Select
    ActiveSheet.Paste
    Range("S2").Select
    Selection.Cut
    Range("R4").Select
    ActiveSheet.Paste
    Range("X2").Select
    Selection.Cut
    Range("W4").Select
    ActiveSheet.Paste
    Range("AC2").Select
    Selection.Cut
    Range("AB4").Select
    ActiveSheet.Paste
    Range("AH2").Select
    Selection.Cut
    Range("AG4").Select
    ActiveSheet.Paste
    Range("AM2").Select
    Selection.Cut
    Range("AL4").Select
    ActiveSheet.Paste
    Range("AR2").Select
    Selection.Cut
    Range("AQ4").Select
    ActiveSheet.Paste
    Range("AW2").Select
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    Range("BB2").Select
    Selection.Cut
    Range("BA4").Select
    ActiveSheet.Paste
    Range("BG2").Select
    Selection.Cut
    Range("BF4").Select
    ActiveSheet.Paste
    Range("BL2").Select
    Selection.Cut
    Range("BK4").Select
    ActiveSheet.Paste
    Range("BQ2").Select
    Selection.Cut
    Range("BP4").Select
    ActiveSheet.Paste
    Range("BV2").Select
    Selection.Cut
    Range("BU4").Select
    ActiveSheet.Paste
    Range("CA2").Select
    Selection.Cut
    Range("BZ4").Select
    ActiveSheet.Paste
    Range("CF2").Select
    Selection.Cut
    Range("CE4").Select
    ActiveSheet.Paste
    Range("CK2").Select
    Selection.Cut
    Range("CJ4").Select
    ActiveSheet.Paste
    Range("CP2").Select
    Selection.Cut
    Range("CO4").Select
    ActiveSheet.Paste
    Range("CU2").Select
    Selection.Cut
    Range("CT4").Select
    ActiveSheet.Paste
    Range("CZ2").Select
    Selection.Cut
    Range("CY4").Select
    ActiveSheet.Paste
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range("C4").Select
    Selection.Copy
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("C109").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("C5:C109").Select
    Range("C109").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("H4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("H109").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("H5:H109").Select
    Range("H109").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("M109").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("M5:M109").Select
    Range("M109").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R4").Select
    Columns("R:R").ColumnWidth = 31.75
    Application.CutCopyMode = False
    Selection.Copy
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("R109").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("R5:R109").Select
    Range("R109").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("W109").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("W5:W109").Select
    Range("W109").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("AB109").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AB5:AB109").Select
    Range("AB109").Activate
    ActiveSheet.Paste
    Range("AB108").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AG109").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AG5:AG109").Select
    Range("AG109").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AL109").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AL5:AL109").Select
    Range("AL109").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AQ109").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AQ5:AQ109").Select
    Range("AQ109").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("AV4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AV109").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AV5:AV109").Select
    Range("AV109").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("BA109").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BA5:BA109").Select
    Range("BA109").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BF109").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BF5:BF109").Select
    Range("BF109").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BK109").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BK5:BK109").Select
    Range("BK109").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BP109").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BP5:BP109").Select
    Range("BP109").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BU109").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BU5:BU109").Select
    Range("BU109").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BZ109").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BZ5:BZ109").Select
    Range("BZ109").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("CE109").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CE5:CE109").Select
    Range("CE109").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CJ109").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CJ5:CJ109").Select
    Range("CJ109").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CO109").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CO5:CO109").Select
    Range("CO109").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CT109").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CT5:CT109").Select
    Range("CT109").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CY4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CY109").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CY5:CY109").Select
    Range("CY109").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CY4").Select
    Application.CutCopyMode = False
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CT110").Select
    ActiveSheet.Paste
    Range("CT109").Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CO110").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CJ110").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CE110").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("BZ110").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BU110").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BP110").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BK110").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BF110").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BA110").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("AV110").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AQ110").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AL110").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AG110").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AB110").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("W110").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("R110").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("M110").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("H110").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("C110").Select
    ActiveSheet.Paste
    Range("B109").Select
    Selection.End(xlToLeft).Select
    Range("A109:B109").Select
    Range("B109").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A4:B109").Select
    Range("B109").Activate
    Selection.Copy
    Range("C18").Select
    Selection.End(xlDown).Select
    Range("A2229:B2229").Select
    Range("B2229").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A110:B2229").Select
    Range("B2229").Activate
    ActiveSheet.Paste
    Range("C2229").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "Target"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Modalidades"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "Respuesta"
    Range("D2:G2").Select
    Selection.Cut Destination:=Range("D3:G3")
    Rows("1:2").Select
    Range("A2").Activate
    Selection.Delete Shift:=xlUp
    Columns("H:H").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 90
    Columns("H:DC").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Variable"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "TRANSPORTE"
    Range("A2").Select
    Selection.Copy
    Range("B2").Select
    Selection.End(xlDown).Select
    Range("A2227").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A3:A2227").Select
    Range("A2227").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Range("A2:H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearFormats
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    Columns("A:A").ColumnWidth = 19.13
    Range("C1").Select
    Selection.Copy
    Range("B1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A2").Select
    Selection.End(xlDown).Select
    Range("A2227").Select
    Selection.End(xlUp).Select
    ActiveWorkbook.Save

    Else
    MsgBox ERROR
End If
End Sub




'==============================================| FIN TRANSPORTE FINANCIEROS | =====================================================================







'==============================================| MEDIOS UP | =====================================================================


Sub MEDIOSUPTOTAL()

Dim NOMBRE As String
Dim ERROR As String
Dim OK As String


ERROR = "LA MACRO SELECCIONADA NO ES VALIDA"
NOMBRE = "MEDIOS U.P."
OK = "OK"


If ActiveSheet.Name = NOMBRE Then
   


'
' Macro2 Macro
'
' Acceso directo: CTRL+q
'
    CONTROLADOR(0909)
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B5").Select
    Selection.Cut
    Range("A6").Select
    ActiveSheet.Paste
    Range("B13").Select
    Selection.Cut
    Range("A14").Select
    ActiveSheet.Paste
    Range("C20").Select
    Selection.End(xlDown).Select
    Range("B25").Select
    Selection.Cut
    Range("A26").Select
    ActiveSheet.Paste
    Range("B28").Select
    Selection.Cut
    Range("A29").Select
    ActiveSheet.Paste
    Range("C48").Select
    Selection.End(xlDown).Select
    Range("B69").Select
    Selection.Cut
    Range("A70").Select
    ActiveSheet.Paste
    Range("A60").Select
    Selection.End(xlUp).Select
    Range("B30").Select
    Selection.Cut
    Range("A31").Select
    ActiveSheet.Paste
    Range("C32").Select
    Selection.End(xlDown).Select
    Range("C76").Select
    Selection.End(xlDown).Select
    Range("B160").Select
    Selection.Cut
    Range("A161").Select
    ActiveSheet.Paste
    Range("C161").Select
    Selection.End(xlDown).Select
    Range("B287").Select
    Selection.Cut
    Range("A288").Select
    ActiveSheet.Paste
    Range("C288").Select
    Selection.End(xlDown).Select
    Range("B308").Select
    Selection.Cut
    Range("A309").Select
    ActiveSheet.Paste
    Range("B315").Select
    Selection.Cut
    Range("A316").Select
    ActiveSheet.Paste
    Range("B325").Select
    Selection.Cut
    Range("A326").Select
    ActiveSheet.Paste
    Range("B332").Select
    Selection.Cut
    Range("A333").Select
    ActiveSheet.Paste
    Range("B340").Select
    Selection.Cut
    Range("A341").Select
    ActiveSheet.Paste
    Range("C342").Select
    Selection.End(xlDown).Select
    Range("B353").Select
    Selection.Cut
    Range("A354").Select
    ActiveSheet.Paste
    Range("B365").Select
    Selection.Cut
    Range("A366").Select
    ActiveSheet.Paste
    Range("B378").Select
    Selection.Cut
    Range("A379").Select
    ActiveSheet.Paste
    Range("B388").Select
    Selection.Cut
    Range("A389").Select
    ActiveSheet.Paste
    Range("B396").Select
    Selection.Cut
    Range("A397").Select
    ActiveSheet.Paste
    Range("B404").Select
    Selection.Cut
    Range("A405").Select
    ActiveSheet.Paste
    Range("B412").Select
    Selection.Cut
    Range("A413").Select
    ActiveSheet.Paste
    Range("B423").Select
    Selection.Cut
    Range("A424").Select
    ActiveSheet.Paste
    Range("B434").Select
    ActiveWindow.ScrollRow = 410
    ActiveWindow.ScrollRow = 408
    ActiveWindow.ScrollRow = 394
    ActiveWindow.ScrollRow = 336
    ActiveWindow.ScrollRow = 225
    ActiveWindow.ScrollRow = 224
    ActiveWindow.ScrollRow = 210
    ActiveWindow.ScrollRow = 106
    ActiveWindow.ScrollRow = 85
    ActiveWindow.ScrollRow = 83
    ActiveWindow.ScrollRow = 82
    ActiveWindow.ScrollRow = 76
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 29
    ActiveWindow.ScrollRow = 13
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 1
    Columns("A:A").ColumnWidth = 21.5
    Rows("4:5").Select
    Range("A5").Activate
    Selection.Delete Shift:=xlUp
    Range("11:11,23:23").Select
    Range("A23").Activate
    Selection.Delete Shift:=xlUp
    Range("C16").Select
    ActiveWindow.SmallScroll Down:=18
    Range("24:24,26:26").Select
    Range("A26").Activate
    Selection.Delete Shift:=xlUp
    Rows("63:63").Select
    Selection.Delete Shift:=xlUp
    Range("A63").Select
    Selection.End(xlDown).Select
    Rows("153:153").Select
    Selection.Delete Shift:=xlUp
    Range("A153").Select
    Selection.End(xlDown).Select
    Rows("279:279").Select
    Selection.Delete Shift:=xlUp
    Range("A279").Select
    Selection.End(xlDown).Select
    Rows("299:299").Select
    Selection.Delete Shift:=xlUp
    Rows("305:305").Select
    Selection.Delete Shift:=xlUp
    Range("A305").Select
    Selection.End(xlDown).Select
    Rows("314:314").Select
    Selection.Delete Shift:=xlUp
    Range("B315").Select
    Selection.End(xlDown).Select
    Range("320:320,328:328").Select
    Range("A328").Activate
    Selection.Delete Shift:=xlUp
    Rows("339:339").Select
    Selection.Delete Shift:=xlUp
    Range("A339").Select
    Selection.End(xlDown).Select
    Rows("350:350").Select
    Selection.Delete Shift:=xlUp
    Rows("362:362").Select
    Selection.Delete Shift:=xlUp
    Rows("371:371").Select
    Selection.Delete Shift:=xlUp
    Range("378:378,386:386").Select
    Range("A386").Activate
    Selection.Delete Shift:=xlUp
    Rows("392:392").Select
    Selection.Delete Shift:=xlUp
    Rows("402:402").Select
    Selection.Delete Shift:=xlUp
    Range("B405").Select
    ActiveWindow.ScrollRow = 392
    ActiveWindow.ScrollRow = 391
    ActiveWindow.ScrollRow = 370
    ActiveWindow.ScrollRow = 261
    ActiveWindow.ScrollRow = 232
    ActiveWindow.ScrollRow = 195
    ActiveWindow.ScrollRow = 114
    ActiveWindow.ScrollRow = 83
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    Range("A4").Select
    Columns("A:A").ColumnWidth = 66.25
    Selection.Copy
    Range("A5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A5:A10").Select
    ActiveSheet.Paste
    Range("A11").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A12").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A12:A21").Select
    ActiveSheet.Paste
    Range("A22").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A23").Select
    ActiveSheet.Paste
    Range("A25").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A26").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A26:A62").Select
    ActiveSheet.Paste
    Range("A28").Select
    Selection.End(xlDown).Select
    Range("A63").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A64").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A64:A152").Select
    ActiveSheet.Paste
    Range("A66").Select
    Selection.End(xlDown).Select
    Range("A153").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A154").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A154:A278").Select
    ActiveSheet.Paste
    Range("A156").Select
    Selection.End(xlDown).Select
    Range("A279").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A280").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A280:A298").Select
    ActiveSheet.Paste
    Range("A281").Select
    Selection.End(xlDown).Select
    Range("A299").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A300").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A300:A304").Select
    ActiveSheet.Paste
    Range("A305").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A306").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A306:A313").Select
    ActiveSheet.Paste
    Range("A314").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A315").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A315:A319").Select
    ActiveSheet.Paste
    Range("A320").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A321").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A321:A326").Select
    ActiveSheet.Paste
    Range("A327").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A328").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A328:A338").Select
    ActiveSheet.Paste
    Range("A339").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A340").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A340:A349").Select
    ActiveSheet.Paste
    Range("A350").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A351").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A351:A361").Select
    ActiveSheet.Paste
    Range("A362").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A363").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A363:A370").Select
    ActiveSheet.Paste
    Range("A371").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A372").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A372:A377").Select
    ActiveSheet.Paste
    Range("A378").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A379").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A379:A384").Select
    ActiveSheet.Paste
    Range("A385").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A386").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A386:A391").Select
    ActiveSheet.Paste
    Range("A392").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A393").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A393:A401").Select
    ActiveSheet.Paste
    Range("A402").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A403:A407").Select
    ActiveSheet.Paste
    Range("B403").Select
    Application.CutCopyMode = False
    Rows("402:402").RowHeight = 22.5
    Range("A401").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A2").Select
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 1
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    Columns("R:R").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("W:W").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    Columns("AB:AB").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AG:AG").Select
    Columns("AG:AG").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    Columns("AL:AL").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AQ:AQ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AV:AV").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    Columns("BA:BA").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BF:BF").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    Columns("BK:BK").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BP:BP").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BU:BU").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    Columns("BZ:BZ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CE:CE").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    Columns("CJ:CJ").Select
    Columns("CJ:CJ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CO:CO").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    Columns("CT:CT").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CY:CY").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 90
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 1
    Range("D1").Select
    Selection.Cut
    Range("D1").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("C4").Select
    ActiveSheet.Paste
    Range("I2").Select
    Selection.Cut
    Range("H4").Select
    ActiveSheet.Paste
    Range("N2").Select
    Selection.Cut
    Range("M4").Select
    ActiveSheet.Paste
    Range("S2").Select
    Selection.Cut
    Range("R4").Select
    ActiveSheet.Paste
    Range("X2").Select
    Selection.Cut
    Range("W4").Select
    ActiveSheet.Paste
    Range("AC2").Select
    Selection.Cut
    Range("AC2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("AB4").Select
    ActiveSheet.Paste
    Range("AH2").Select
    Selection.Cut
    Range("AG4").Select
    ActiveSheet.Paste
    Range("AM2").Select
    Selection.Cut
    Range("AL4").Select
    ActiveSheet.Paste
    Range("AR2").Select
    Selection.Cut
    Range("AQ4").Select
    ActiveSheet.Paste
    Range("AW2").Select
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    Range("BB2").Select
    Selection.Cut
    Range("BA4").Select
    ActiveSheet.Paste
    Range("BG2").Select
    Selection.Cut
    Range("BF4").Select
    ActiveSheet.Paste
    Range("BL2").Select
    Selection.Cut
    Range("BK4").Select
    ActiveSheet.Paste
    Range("BQ2").Select
    Selection.Cut
    Range("BP4").Select
    ActiveSheet.Paste
    Range("BV2").Select
    Selection.Cut
    Range("BU4").Select
    ActiveSheet.Paste
    Range("CA2").Select
    Selection.Cut
    Range("BZ4").Select
    ActiveSheet.Paste
    Range("CF2").Select
    Selection.Cut
    Range("CE4").Select
    ActiveSheet.Paste
    Range("CK2").Select
    Selection.Cut
    Range("CJ4").Select
    ActiveSheet.Paste
    Range("CP2").Select
    Selection.Cut
    Range("CO4").Select
    ActiveSheet.Paste
    Range("CU2").Select
    Selection.Cut
    Range("CT4").Select
    ActiveSheet.Paste
    Range("CZ2").Select
    Selection.Cut
    Range("CY4").Select
    ActiveSheet.Paste
    Range("C4").Select
    Selection.Copy
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("C407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("C5:C407").Select
    Range("C407").Activate
    ActiveSheet.Paste
    Range("D407").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("C18").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("C4").Select
    Application.CutCopyMode = False
    Range("H4").Select
    Selection.Cut
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("H407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("H5:H407").Select
    Range("H407").Activate
    Selection.End(xlUp).Select
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("H407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("H5:H407").Select
    Range("H407").Activate
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("H407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("H5:H407").Select
    Range("H407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("M407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("M5:M407").Select
    Range("M407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("R407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("R5:R407").Select
    Range("R407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("W407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("W5:W407").Select
    Range("W407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W3").Select
    Application.CutCopyMode = False
    Range("AB4").Select
    Selection.Copy
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("AB407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AB5:AB407").Select
    Range("AB407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Range("AG4").Select
    Selection.Copy
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AG407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AG5:AG407").Select
    Range("AG407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Range("AL4").Select
    Selection.Copy
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AL407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AL5:AL407").Select
    Range("AL407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL3").Select
    Application.CutCopyMode = False
    Range("AQ4").Select
    Selection.Copy
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AQ407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AQ5:AQ407").Select
    Range("AQ407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AV407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AV5:AV407").Select
    Range("AV407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("BA407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BA5:BA407").Select
    Range("BA407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BF407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BF5:BF407").Select
    Range("BF407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BK407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BK5:BK407").Select
    Range("BK407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BP407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BP5:BP407").Select
    Range("BP407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BU407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BU5:BU407").Select
    Range("BU407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BZ407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BZ5:BZ407").Select
    Range("BZ407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("CE407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CE5:CE407").Select
    Range("CE407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CJ407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CJ5:CJ407").Select
    Range("CJ407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CO407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CO5:CO407").Select
    Range("CO407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CT407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CT5:CT407").Select
    Range("CT407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CY4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CY407").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CY5:CY407").Select
    Range("CY407").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("D4").Select
    Selection.End(xlToRight).Select
    Range("CY4").Select
    Application.CutCopyMode = False
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CT408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CO408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CJ408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CE408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("BZ408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BU408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BP408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BK408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BF5").Select
    Selection.End(xlDown).Select
    Range("BF408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BA408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("AV408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AQ408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AL408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AG408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AB408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("W408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("R408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("M408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("H408").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("C408").Select
    ActiveSheet.Paste
    Range("B407").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A4:B407").Select
    Range("B407").Activate
    Selection.Copy
    Range("C407").Select
    Selection.End(xlDown).Select
    Range("A8487:B8487").Select
    Range("B8487").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A408:B8487").Select
    Range("B8487").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Columns("H:H").Select
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 90
    Columns("H:DC").Select
    Selection.Delete Shift:=xlToLeft
    Range("G5").Select
    ActiveWindow.ScrollColumn = 1
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Target"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Respuesta"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Modalidades"
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Variable"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "MEDIOS"
    Range("A2").Select
    Selection.Copy
    Range("B2").Select
    Selection.End(xlDown).Select
    Range("A8485").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A3:A8485").Select
    Range("A8485").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("C1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearFormats
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.349986266670736
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.349986266670736
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.349986266670736
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.349986266670736
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.349986266670736
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.349986266670736
        .Weight = xlThin
    End With
    Range("B14").Select
    Columns("B:B").ColumnWidth = 54.38
    Columns("C:C").ColumnWidth = 23.25
    Range("A7").Select
    Selection.End(xlDown).Select
    Range("A8485").Select
    Selection.End(xlUp).Select
    ActiveWorkbook.Save
    CLEARCONTROLER(1010)

    Else
    MsgBox ERROR
End If
End Sub

'==============================================| FIN MEDIOS UP | =====================================================================


'==============================================| MEDIOS U30 | =====================================================================

Sub MEDIOSU30()

Dim NOMBRE As String
Dim ERROR As String
Dim OK As String


ERROR = "LA MACRO SELECCIONADA NO ES VALIDA"
NOMBRE = "MEDIOS U.30"
OK = "OK"


If ActiveSheet.Name = NOMBRE Then
  

'
' Macro3 Macro
'
' Acceso directo: CTRL+w
'
    CONTROLADOR(0909)
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.ColumnWidth = 44.13
    Range("B5").Select
    Selection.Cut
    Range("A6").Select
    ActiveSheet.Paste
    Range("B13").Select
    Selection.Cut
    Range("A14").Select
    ActiveSheet.Paste
    Range("B25").Select
    Selection.Cut
    Range("A26").Select
    ActiveSheet.Paste
    Range("B28").Select
    Selection.Cut
    Range("A29").Select
    ActiveSheet.Paste
    Range("B30").Select
    Selection.Cut
    Range("A31").Select
    ActiveSheet.Paste
    Range("C31").Select
    Selection.End(xlDown).Select
    Range("B79").Select
    Selection.Cut
    Range("A80").Select
    ActiveSheet.Paste
    Range("C132").Select
    Selection.End(xlDown).Select
    Range("B191").Select
    Selection.Cut
    Range("A192").Select
    ActiveSheet.Paste
    Range("C193").Select
    Selection.End(xlDown).Select
    Range("B217").Select
    Selection.Cut
    Range("A218").Select
    ActiveSheet.Paste
    Range("C218").Select
    Selection.End(xlDown).Select
    Range("B343").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Rows("4:4").Select
    Selection.Delete Shift:=xlUp
    Rows("4:4").Select
    Selection.Delete Shift:=xlUp
    Rows("11:11").Select
    Selection.Delete Shift:=xlUp
    Rows("22:22").Select
    Selection.Delete Shift:=xlUp
    Rows("24:24").Select
    Selection.Delete Shift:=xlUp
    Rows("25:25").Select
    Selection.Delete Shift:=xlUp
    Range("A57").Select
    Selection.End(xlDown).Select
    Rows("73:73").Select
    Selection.Delete Shift:=xlUp
    Range("A88").Select
    Selection.End(xlDown).Select
    Rows("184:184").Select
    Selection.Delete Shift:=xlUp
    Range("A184").Select
    Selection.End(xlDown).Select
    Rows("209:209").Select
    Selection.Delete Shift:=xlUp
    Range("A209").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("B209").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A4").Select
    Selection.Copy
    Range("A5:A10").Select
    ActiveSheet.Paste
    Range("A11").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A12:A21").Select
    ActiveSheet.Paste
    Range("A22").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A23").Select
    ActiveSheet.Paste
    Range("A25").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A26").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A26:A72").Select
    ActiveSheet.Paste
    Range("A27").Select
    Selection.End(xlDown).Select
    Range("A73").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A74").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A74:A183").Select
    ActiveSheet.Paste
    Range("A77").Select
    Selection.End(xlDown).Select
    Range("A184").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A185").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A185:A208").Select
    ActiveSheet.Paste
    Range("A187").Select
    Selection.End(xlDown).Select
    Range("A209").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A210").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("B210").Select
    Selection.End(xlDown).Select
    Range("A334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A210:A334").Select
    Range("A334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    Columns("R:R").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("W:W").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AB:AB").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AG:AG").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AL:AL").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AQ:AQ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AV:AV").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BA:BA").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BF:BF").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BK:BK").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BP:BP").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BU:BU").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BZ:BZ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CE:CE").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CJ:CJ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CO:CO").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CT:CT").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CY:CY").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D1").Select
    Selection.Cut
    Range("C4").Select
    ActiveSheet.Paste
    Range("I2").Select
    Selection.Cut
    Range("H4").Select
    ActiveSheet.Paste
    Range("N2").Select
    Selection.Cut
    Range("M4").Select
    ActiveSheet.Paste
    Range("S2").Select
    Selection.Cut
    Range("R4").Select
    ActiveSheet.Paste
    Range("X2").Select
    Selection.Cut
    Range("W4").Select
    ActiveSheet.Paste
    Range("AC2").Select
    Selection.Cut
    Range("AB4").Select
    ActiveSheet.Paste
    Range("AH2").Select
    Selection.Cut
    Range("AG4").Select
    ActiveSheet.Paste
    Range("AM2").Select
    Selection.Cut
    Range("AL4").Select
    ActiveSheet.Paste
    Range("AR2").Select
    Selection.Cut
    Range("AQ4").Select
    ActiveSheet.Paste
    Range("AW2").Select
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    Range("BB2").Select
    Selection.Cut
    Range("BA4").Select
    ActiveSheet.Paste
    Range("BG2").Select
    Selection.Cut
    Range("BF4").Select
    ActiveSheet.Paste
    Range("BL2").Select
    Selection.Cut
    Selection.End(xlDown).Select
    Range("BK4").Select
    ActiveSheet.Paste
    Range("BQ2").Select
    Selection.Cut
    Range("BP4").Select
    ActiveSheet.Paste
    Range("BV2").Select
    Selection.Cut
    Range("BU4").Select
    ActiveSheet.Paste
    Range("CA2").Select
    Selection.Cut
    Range("BZ4").Select
    ActiveSheet.Paste
    Range("CF2").Select
    Selection.Cut
    Range("CE4").Select
    ActiveSheet.Paste
    Range("CK2").Select
    Selection.Cut
    Range("CJ4").Select
    ActiveSheet.Paste
    Range("CP2").Select
    Selection.Cut
    Range("CO4").Select
    ActiveSheet.Paste
    Range("CU2").Select
    Selection.Cut
    Range("CT4").Select
    ActiveSheet.Paste
    Range("CZ2").Select
    Selection.Cut
    Range("CY4").Select
    ActiveSheet.Paste
    Range("CP1").Select
    Selection.End(xlToLeft).Select
    Range("C4").Select
    Selection.Copy
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("C334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("C5:C334").Select
    Range("C334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("H334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("H5:H334").Select
    Range("H334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("M334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("M5:M334").Select
    Range("M334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("R334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("R5:R334").Select
    Range("R334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("W334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("W5:W334").Select
    Range("W334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("AB334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AB5:AB334").Select
    Range("AB334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AG334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AG5:AG334").Select
    Range("AG334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AL334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AL5:AL334").Select
    Range("AL334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AQ334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AQ5:AQ334").Select
    Range("AQ334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AV334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AV5:AV334").Select
    Range("AV334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("BA334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BA5:BA334").Select
    Range("BA334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BF334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BF5:BF334").Select
    Range("BF334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BK334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BK5:BK334").Select
    Range("BK334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BP334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BP5:BP334").Select
    Range("BP334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BU334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BU5:BU334").Select
    Range("BU334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BZ334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BZ5:BZ334").Select
    Range("BZ334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("CE334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CE5:CE334").Select
    Range("CE334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CJ334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CJ5:CJ334").Select
    Range("CJ334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CO334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CO5:CO334").Select
    Range("CO334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CT334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CT5:CT334").Select
    Range("CT334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CY4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CY334").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CY5:CY334").Select
    Range("CY334").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("DF4").Select
    Application.CutCopyMode = False
    Range("CY4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CT335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CO335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CJ335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CE335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("BZ335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BU335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BP335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BK335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BF335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BA335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("AV335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AQ335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AL335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AG335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AB335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("W335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("R335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("S4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("R4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("M335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("H335").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("C335").Select
    ActiveSheet.Paste
    Range("A334:B334").Select
    Range("B334").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A4:B334").Select
    Range("B334").Activate
    Selection.Copy
    Range("C335").Select
    Selection.End(xlDown).Select
    Range("A6954:B6954").Select
    Range("B6954").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A335:B6954").Select
    Range("B6954").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Columns("H:H").Select
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 90
    Columns("H:DC").Select
    Selection.Delete Shift:=xlToLeft
    Rows("3:3").Select
    Selection.Delete Shift:=xlUp
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Target"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Respuesta"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Modalidades"
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearFormats
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Variable"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "MEDIOS"
    Range("A2").Select
    Selection.Copy
    Range("B2").Select
    Selection.End(xlDown).Select
    Range("A6952").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A3:A6952").Select
    Range("A6952").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Range("C1").Select
    Selection.Copy
    Range("B1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    Range("B10").Select
    Columns("C:C").EntireColumn.AutoFit
    Columns("B:B").ColumnWidth = 39.5
    Range("A18").Select
    Selection.End(xlDown).Select
    Range("A6952").Select
    Selection.End(xlUp).Select
    ActiveWorkbook.Save
    CLEARCONTROLER(1111)

    Else
    MsgBox ERROR
End If
End Sub

'==============================================| FIN MEDIOS U30 | =====================================================================



'==============================================| MEDIOS DIA AYER | =====================================================================


Sub MEDIOSDIAAYER()



Dim NOMBRE As String
Dim ERROR As String
Dim OK As String


ERROR = "LA MACRO SELECCIONADA NO ES VALIDA"
NOMBRE = "MEDIO D.AYER"
OK = "OK"


If ActiveSheet.Name = NOMBRE Then


'
' MEDIOSDIAAYER Macro
'
' Acceso directo: CTRL+t
'
    CONTROLADOR(696)
    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("L:L").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("Q:Q").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    Columns("V:V").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AA:AA").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    Columns("AF:AF").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AK:AK").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 31
    Columns("AP:AP").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AU:AU").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    Columns("AZ:AZ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BE:BE").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 51
    Columns("BJ:BJ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BO:BO").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    Columns("BT:BT").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BY:BY").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    Columns("CD:CD").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CI:CI").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    Columns("CN:CN").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CS:CS").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    Columns("CX:CX").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("DA23").Select
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("H1").Select
    Selection.Cut
    Range("G4").Select
    ActiveSheet.Paste
    Range("M2").Select
    Selection.Cut
    Range("L4").Select
    ActiveSheet.Paste
    Range("R2").Select
    Selection.Cut
    Range("Q4").Select
    ActiveSheet.Paste
    Range("W2").Select
    Selection.Cut
    Range("V4").Select
    ActiveSheet.Paste
    Range("AB2").Select
    Selection.Cut
    Range("AA4").Select
    ActiveSheet.Paste
    Range("AG2").Select
    Selection.Cut
    Range("AF4").Select
    ActiveSheet.Paste
    Range("AL2").Select
    Selection.Cut
    Range("AK4").Select
    ActiveSheet.Paste
    Range("AQ2").Select
    Selection.Cut
    Range("AP4").Select
    ActiveSheet.Paste
    Range("AV2").Select
    Selection.Cut
    Range("AU4").Select
    ActiveSheet.Paste
    Range("BA2").Select
    Selection.Cut
    Range("AZ4").Select
    ActiveSheet.Paste
    Range("BF2").Select
    Selection.Cut
    Range("BE4").Select
    ActiveSheet.Paste
    Range("BK2").Select
    Selection.Cut
    Range("BJ4").Select
    ActiveSheet.Paste
    Range("BP2").Select
    Selection.Cut
    Range("BO4").Select
    ActiveSheet.Paste
    Range("BU2").Select
    Selection.Cut
    Range("BT4").Select
    ActiveSheet.Paste
    Range("BZ2").Select
    Selection.Cut
    Range("BY4").Select
    ActiveSheet.Paste
    Range("CE2").Select
    Selection.Cut
    Range("CD4").Select
    ActiveSheet.Paste
    Range("CJ2").Select
    Selection.Cut
    Selection.End(xlDown).Select
    Range("CJ86").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CI4").Select
    ActiveSheet.Paste
    Range("CO2").Select
    Selection.Cut
    Range("CN4").Select
    ActiveSheet.Paste
    Range("CT2").Select
    Selection.Cut
    Range("CS4").Select
    ActiveSheet.Paste
    Range("CY2").Select
    Selection.Cut
    Range("CX4").Select
    ActiveSheet.Paste
    Selection.End(xlToLeft).Select
    Range("C1").Select
    Selection.Cut
    Range("B4").Select
    ActiveSheet.Paste
    Range("B4").Select
    Selection.Copy
    Range("A5").Select
    Selection.End(xlDown).Select
    Range("B120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("B5:B120").Select
    Range("B120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H2").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("G4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("G120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("G5:G120").Select
    Range("G120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("L4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G4").Select
    Selection.End(xlDown).Select
    Range("L120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("L5:L120").Select
    Range("L120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("Q4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("L4").Select
    Selection.End(xlDown).Select
    Range("Q120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("Q5:Q120").Select
    Range("Q120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("V4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("Q4").Select
    Selection.End(xlDown).Select
    Range("V120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("V5:V120").Select
    Range("V120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("AA4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("V4").Select
    Selection.End(xlDown).Select
    Range("AA120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AA5:AA120").Select
    Range("AA120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AF4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AA4").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("AF120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AF5:AF120").Select
    Range("AF120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("AK4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AF4").Select
    Selection.End(xlDown).Select
    Range("AK120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AK5:AK120").Select
    Range("AK120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AP4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AK4").Select
    Selection.End(xlDown).Select
    Range("AP120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AP5:AP120").Select
    Range("AP120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AU4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AP4").Select
    Selection.End(xlDown).Select
    Range("AU120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AU5:AU120").Select
    Range("AU120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("AZ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AU4").Select
    Selection.End(xlDown).Select
    Range("AZ120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AZ5:AZ120").Select
    Range("AZ120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BE4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AZ4").Select
    Selection.End(xlDown).Select
    Range("BE120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BE5:BE120").Select
    Range("BE120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BJ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BE4").Select
    Selection.End(xlDown).Select
    Range("BJ120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BJ5:BJ120").Select
    Range("BJ120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BO4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BJ4").Select
    Selection.End(xlDown).Select
    Range("BO120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BO5:BO120").Select
    Range("BO120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BT4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BO4").Select
    Selection.End(xlDown).Select
    Range("BT120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BT5:BT120").Select
    Range("BT120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BY4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BT4").Select
    Selection.End(xlDown).Select
    Range("BY120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BY5:BY120").Select
    Range("BY120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CD4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BY4").Select
    Selection.End(xlDown).Select
    Range("CD120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CD5:CD120").Select
    Range("CD120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CI4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CD4").Select
    Selection.End(xlDown).Select
    Range("CI120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CI5:CI120").Select
    Range("CI120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CN4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CI4").Select
    Selection.End(xlDown).Select
    Range("CN120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CN5:CN120").Select
    Range("CN120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CS4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CN4").Select
    Selection.End(xlDown).Select
    Range("CS120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CS5:CS120").Select
    Range("CS120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CX4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CS4").Select
    Selection.End(xlDown).Select
    Range("CX120").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CX5:CX120").Select
    Range("CX120").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CX4").Select
    Application.CutCopyMode = False
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("CX4").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Rows("4:4").Select
    Selection.Delete Shift:=xlUp
    Range("G4").Select
    Selection.End(xlToRight).Select
    Range("CX4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CS4").Select
    Selection.End(xlDown).Select
    Range("CS120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CS4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CN4").Select
    Selection.End(xlDown).Select
    Range("CN120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CN4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CI4").Select
    Selection.End(xlDown).Select
    Range("CI120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CI4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CD4").Select
    Selection.End(xlDown).Select
    Range("CD120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CD4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BY4").Select
    Selection.End(xlDown).Select
    Range("BY120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BY4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BT4").Select
    Selection.End(xlDown).Select
    Range("BT120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BT4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BO4").Select
    Selection.End(xlDown).Select
    Range("BO120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BO4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BJ4").Select
    Selection.End(xlDown).Select
    Range("BJ120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BJ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BE6").Select
    Selection.End(xlDown).Select
    Range("BE120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BE4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AZ4").Select
    Selection.End(xlDown).Select
    Range("AZ120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AZ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AU4").Select
    Selection.End(xlDown).Select
    Range("AU120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AU4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AP4").Select
    Selection.End(xlDown).Select
    Range("AP120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AP4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AK4").Select
    Selection.End(xlDown).Select
    Range("AK120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AK4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AF4").Select
    Selection.End(xlDown).Select
    Range("AF120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AF4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AA4").Select
    Selection.End(xlDown).Select
    Range("AA120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AA4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("V4").Select
    Selection.End(xlDown).Select
    Range("V120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("V4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("Q4").Select
    Selection.End(xlDown).Select
    Range("Q120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("Q4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("L4").Select
    Selection.End(xlDown).Select
    Range("L120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("L4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("G4").Select
    Selection.End(xlDown).Select
    Range("G120").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("G4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("B120").Select
    ActiveSheet.Paste
    Range("A119").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A4:A119").Select
    Range("A119").Activate
    Selection.Copy
    Range("B120").Select
    Selection.End(xlDown).Select
    Range("A2439").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A120:A2439").Select
    Range("A2439").Activate
    ActiveSheet.Paste
    Range("A2439").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Columns("G:G").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 89
    Columns("G:DB").Select
    Selection.Delete Shift:=xlToLeft
    Rows("3:3").Select
    Selection.Delete Shift:=xlUp
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Variable"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Modalidades"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Respuesta"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Target"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "Television el dia de ayer"
    Range("B2").Select
    Selection.Copy
    Range("C2").Select
    Selection.End(xlDown).Select
    Range("B2437").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("B3:B2437").Select
    Range("B2437").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "MEDIOS"
    Range("A2").Select
    Selection.Copy
    Range("B2").Select
    Selection.End(xlDown).Select
    Range("A2437").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A3:A2437").Select
    Range("A2437").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Range("C1").Select
    Selection.Copy
    Range("B1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A2:H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearFormats
    Columns("C:C").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
    Range("A2:H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    Range("B9").Select
    ActiveWindow.SmallScroll Down:=-15
    ActiveWorkbook.Save
   CLEARCONTROLER(1212)
    Else
    MsgBox ERROR
End If
End Sub


'==============================================| FIN MEDIOS DIA AYER | =====================================================================




'=========================================| PERFIL_MODULOS|============================================'




Private  Sub PERFIL_ADD_TOTAL()

'
' TOTAL Macro
'

'
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("A5").Select
    ActiveSheet.Paste
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("B2:E2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("B1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range("B1:E1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.Cut
    Application.CutCopyMode = False
    Selection.Copy
    Range("B1").Select
    ActiveSheet.Paste
    Range("B1").Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Range("B2").Select
    Selection.EntireColumn.Insert
    Range("C2").Select
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Columns("B:B").ColumnWidth = 17.14
    Columns("C:C").ColumnWidth = 17.14
    Columns("D:D").ColumnWidth = 17.14
    Columns("E:E").ColumnWidth = 17.14
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("F3:I3").Select
    Selection.Copy
    Range("B2").Select
    ActiveSheet.Paste
    Range("C3").Select
    Application.CutCopyMode = False
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("B4:E4").Select
    Selection.Copy
    Range("C4").Select
    Selection.End(xlToRight).Select
    Range("F4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("F4").Select
    Selection.End(xlToRight).Select
    Range("XFB5").Select
    Selection.End(xlToLeft).Select
    Range("B7").Select
    Selection.End(xlToRight).Select
    Range("F6").Select
    Selection.End(xlToRight).Select
    Range("CG4").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range("F4:CG4").Select
    Range("CG4").Activate
    ActiveSheet.Paste
    Range("CF4").Select
    Selection.End(xlToLeft).Select
    Range("C4").Select
    Application.CutCopyMode = False
    Range("F6:I6").Select
    Application.CutCopyMode = False
    Range("B1").Select
End Sub

Private  Sub PERFIL_CLEAR_TOTAL()
   Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Rows("2:62").Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select

End Sub


'==============================================| FIN PERFIL_MODULOS |================================================================

'==============================================|     INTERNET   |==========================================================

Sub INTERNET()
Dim NOMBRE As String
Dim ERROR As String
Dim OK As String


ERROR = "LA MACRO SELECCIONADA NO ES VALIDA"
NOMBRE = "INTERNET"
OK = "OK"


If ActiveSheet.Name = NOMBRE Then
  



'
' Internet Macro
'
' Acceso directo: CTRL+r
'
    CONTROLADOR(0909)
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B5").Select
    Columns("A:A").ColumnWidth = 88.13
    Columns("A:A").ColumnWidth = 66.13
    Selection.Cut
    Range("A6").Select
    ActiveSheet.Paste
    Range("B8").Select
    Selection.Cut
    Range("B8").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("A9").Select
    ActiveSheet.Paste
    Range("B14").Select
    Selection.Cut
    Range("A15").Select
    ActiveSheet.Paste
    Range("B24").Select
    Selection.Cut
    Range("A25").Select
    ActiveSheet.Paste
    Range("C25").Select
    Selection.End(xlDown).Select
    Range("B33").Select
    Selection.Cut
    Range("A34").Select
    ActiveSheet.Paste
    Range("C34").Select
    Selection.End(xlDown).Select
    Range("B36").Select
    Selection.Cut
    Range("A37").Select
    ActiveSheet.Paste
    Range("C37").Select
    Selection.End(xlDown).Select
    Range("B52").Select
    Selection.Cut
    Range("A53").Select
    ActiveSheet.Paste
    Range("B55").Select
    Selection.Cut
    Range("A56").Select
    ActiveSheet.Paste
    Range("C56").Select
    Selection.End(xlDown).Select
    Range("B60").Select
    Selection.Cut
    Range("A61").Select
    ActiveSheet.Paste
    Range("C61").Select
    Selection.End(xlDown).Select
    Range("B68").Select
    Selection.Cut
    Range("A69").Select
    ActiveSheet.Paste
    Range("B76").Select
    Selection.Cut
    Range("A77").Select
    ActiveSheet.Paste
    Range("B81").Select
    Selection.Cut
    Range("A82").Select
    ActiveSheet.Paste
    Range("C82").Select
    Selection.End(xlDown).Select
    Range("B84").Select
    Selection.Cut
    Range("A85").Select
    ActiveSheet.Paste
    Range("B88").Select
    Selection.Cut
    Range("A89").Select
    ActiveSheet.Paste
    Range("B91").Select
    Selection.Cut
    Range("A92").Select
    ActiveSheet.Paste
    Range("B96").Select
    Selection.Cut
    Range("A97").Select
    ActiveSheet.Paste
    Range("B99").Select
    Selection.Cut
    Range("A100").Select
    ActiveSheet.Paste
    Range("B110").Select
    Selection.Cut
    Range("A111").Select
    ActiveSheet.Paste
    Range("B114").Select
    Selection.Cut
    Range("A115").Select
    ActiveSheet.Paste
    Range("C116").Select
    Selection.End(xlDown).Select
    Range("B119").Select
    Selection.Cut
    Range("A120").Select
    ActiveSheet.Paste
    Range("B122").Select
    Selection.End(xlDown).Select
    Range("A196").Select
    Selection.End(xlUp).Select
    Range("B134").Select
    Selection.Cut
    Range("A135").Select
    ActiveSheet.Paste
    Range("B143").Select
    Selection.Cut
    Range("A144").Select
    ActiveSheet.Paste
    Range("B149").Select
    Selection.Cut
    Range("A150").Select
    ActiveSheet.Paste
    Range("B157").Select
    Selection.Cut
    Range("A158").Select
    ActiveSheet.Paste
    Range("B165").Select
    Selection.Cut
    Range("A166").Select
    ActiveSheet.Paste
    Range("B173").Select
    Selection.Cut
    Range("A174").Select
    ActiveSheet.Paste
    Range("B181").Select
    Selection.Cut
    Range("A182").Select
    ActiveSheet.Paste
    Range("B189").Select
    Selection.Cut
    Range("A190").Select
    ActiveSheet.Paste
    Range("B190").Select
    ActiveWindow.ScrollRow = 177
    ActiveWindow.ScrollRow = 176
    ActiveWindow.ScrollRow = 174
    ActiveWindow.ScrollRow = 166
    ActiveWindow.ScrollRow = 160
    ActiveWindow.ScrollRow = 44
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 21
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 1
    Range("A6").Select
    Selection.Copy
    Range("A7:A8").Select
    ActiveSheet.Paste
    Range("A9").Select
    Application.CutCopyMode = False
    Selection.Cut
    Application.CutCopyMode = False
    Selection.Copy
    Range("A10:A14").Select
    ActiveSheet.Paste
    Range("A15").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A16").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A16:A24").Select
    ActiveSheet.Paste
    Range("A25").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A26").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A26:A33").Select
    ActiveSheet.Paste
    Range("A34").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A35:A36").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A37").Select
    Selection.Copy
    Range("A38").Select
    Selection.End(xlDown).Select
    Range("A52").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A38:A52").Select
    Range("A52").Activate
    ActiveSheet.Paste
    Range("A53").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A54:A55").Select
    ActiveSheet.Paste
    Range("A56").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A57").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A57:A60").Select
    ActiveSheet.Paste
    Range("A61").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A62").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A62:A68").Select
    ActiveSheet.Paste
    Range("A69").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A70").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A70:A76").Select
    ActiveSheet.Paste
    Range("A77").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A78").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A78:A81").Select
    ActiveSheet.Paste
    Range("A82").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A83").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A83:A84").Select
    ActiveSheet.Paste
    Range("A85").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A86").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A86:A88").Select
    ActiveSheet.Paste
    Range("A89").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A90:A91").Select
    ActiveSheet.Paste
    Range("A92").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A93").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A93:A96").Select
    ActiveSheet.Paste
    Range("A97").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A98:A99").Select
    ActiveSheet.Paste
    Range("A100").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A101").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A101:A110").Select
    ActiveSheet.Paste
    Range("A111").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A112:A114").Select
    ActiveSheet.Paste
    Range("A115").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A116").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A116:A119").Select
    ActiveSheet.Paste
    Range("A120").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A121").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A121:A134").Select
    ActiveSheet.Paste
    Range("A135").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A136").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A136:A143").Select
    ActiveSheet.Paste
    Range("A144").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A145").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A145:A149").Select
    ActiveSheet.Paste
    Range("A150").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.End(xlDown).Select
    Range("A151").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A151:A157").Select
    ActiveSheet.Paste
    Range("A158").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A159").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A159:A165").Select
    ActiveSheet.Paste
    Range("A166").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A167").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A167:A173").Select
    ActiveSheet.Paste
    Range("A174").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A175").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A175:A181").Select
    ActiveSheet.Paste
    Range("A182").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A183").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A183:A189").Select
    ActiveSheet.Paste
    Range("A190").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A191").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A191:A196").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Rows("4:5").Select
    Selection.Delete Shift:=xlUp
    Rows("6:6").Select
    Selection.Delete Shift:=xlUp
    Rows("11:11").Select
    Selection.Delete Shift:=xlUp
    Rows("20:20").Select
    Selection.Delete Shift:=xlUp
    Rows("28:28").Select
    Selection.Delete Shift:=xlUp
    Rows("30:30").Select
    Selection.Delete Shift:=xlUp
    Rows("45:45").Select
    Selection.Delete Shift:=xlUp
    Rows("47:47").Select
    Selection.Delete Shift:=xlUp
    Rows("51:51").Select
    Selection.Delete Shift:=xlUp
    Rows("58:58").Select
    Selection.Delete Shift:=xlUp
    Rows("70:70").Select
    Selection.Delete Shift:=xlUp
    Rows("72:72").Select
    Selection.Delete Shift:=xlUp
    Rows("75:75").Select
    Selection.Delete Shift:=xlUp
    Rows("77:77").Select
    Selection.Delete Shift:=xlUp
    Rows("81:81").Select
    Selection.Delete Shift:=xlUp
    Rows("83:83").Select
    Selection.Delete Shift:=xlUp
    Rows("93:93").Select
    Selection.Delete Shift:=xlUp
    Rows("96:96").Select
    Selection.Delete Shift:=xlUp
    Rows("100:100").Select
    Selection.Delete Shift:=xlUp
    Rows("114:114").Select
    Selection.Delete Shift:=xlUp
    Rows("122:122").Select
    Selection.Delete Shift:=xlUp
    Rows("127:127").Select
    Selection.Delete Shift:=xlUp
    Rows("134:134").Select
    Selection.Delete Shift:=xlUp
    Rows("141:141").Select
    Selection.Delete Shift:=xlUp
    Rows("148:148").Select
    Selection.Delete Shift:=xlUp
    Rows("155:155").Select
    Selection.Delete Shift:=xlUp
    Rows("162:162").Select
    Selection.Delete Shift:=xlUp
    Range("C168").Select
    Selection.End(xlUp).Select
    Rows("65:65").Select
    Selection.Delete Shift:=xlUp
    Range("C65").Select
    Selection.End(xlUp).Select
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    Columns("R:R").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("W:W").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    Columns("AB:AB").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AG:AG").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    Columns("AL:AL").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AQ:AQ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    Columns("AV:AV").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BA:BA").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 45
    Columns("BF:BF").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BK:BK").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 55
    Columns("BP:BP").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BU:BU").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    Columns("BZ:BZ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CE:CE").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    Columns("CJ:CJ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CO:CO").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CT:CT").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 89
    Columns("CY:CY").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 1
    Range("D1").Select
    Selection.Cut
    Range("C4").Select
    ActiveSheet.Paste
    Range("I2").Select
    Selection.Cut
    Range("H4").Select
    ActiveSheet.Paste
    Range("N2").Select
    Selection.Cut
    Selection.End(xlDown).Select
    Range("N168").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M4").Select
    ActiveSheet.Paste
    Range("S2").Select
    Selection.Cut
    Range("R4").Select
    ActiveSheet.Paste
    Range("X2").Select
    Selection.Cut
    Range("W4").Select
    ActiveSheet.Paste
    Range("AC2").Select
    Selection.Cut
    Range("AB4").Select
    ActiveSheet.Paste
    Range("AH2").Select
    Selection.Cut
    Selection.End(xlDown).Select
    Range("AH168").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    ActiveSheet.Paste
    Range("AM2").Select
    Selection.Cut
    Range("AL4").Select
    ActiveSheet.Paste
    Range("AR2").Select
    Selection.Cut
    Range("AQ4").Select
    ActiveSheet.Paste
    Range("AW2").Select
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    Range("BB2").Select
    Selection.Cut
    Range("BA4").Select
    ActiveSheet.Paste
    Range("BG2").Select
    Selection.Cut
    Range("BF4").Select
    ActiveSheet.Paste
    Range("BL2").Select
    Selection.Cut
    Range("BK4").Select
    ActiveSheet.Paste
    Range("BQ2").Select
    Selection.Cut
    Range("BP4").Select
    ActiveSheet.Paste
    Range("BV2").Select
    Selection.Cut
    Range("BU4").Select
    ActiveSheet.Paste
    Range("CA2").Select
    Selection.Cut
    Range("BZ4").Select
    ActiveSheet.Paste
    Range("CF2").Select
    Selection.Cut
    Range("CE4").Select
    ActiveSheet.Paste
    Range("CK2").Select
    Selection.Cut
    Range("CJ4").Select
    ActiveSheet.Paste
    Range("CP2").Select
    Selection.Cut
    Range("CO4").Select
    ActiveSheet.Paste
    Range("CU2").Select
    Selection.Cut
    Range("CT4").Select
    ActiveSheet.Paste
    Range("CZ2").Select
    Selection.Cut
    Range("CY4").Select
    ActiveSheet.Paste
    Range("H4").Select
    Selection.End(xlToLeft).Select
    Range("C4").Select
    Selection.Copy
    Range("B5").Select
    Selection.End(xlDown).Select
    Range("C167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("C5:C167").Select
    Range("C167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("H4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G5").Select
    Selection.End(xlDown).Select
    Range("C9").Select
    Selection.End(xlDown).Select
    Range("H167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("H5:H167").Select
    Range("H167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("M167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("M5:M167").Select
    Range("M167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("R167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("R5:R167").Select
    Range("R167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("W167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("W5:W167").Select
    Range("W167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("AB167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AB5:AB167").Select
    Range("AB167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AG167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AG5:AG167").Select
    Range("AG167").Activate
    ActiveSheet.Paste
    Range("AG166").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AL167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AL5:AL167").Select
    Range("AL167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AQ167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AQ5:AQ167").Select
    Range("AQ167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AV167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AV5:AV167").Select
    Range("AV167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("BA167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BA5:BA167").Select
    Range("BA167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BF167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BF5:BF167").Select
    Range("BF167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BK167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BK5:BK167").Select
    Range("BK167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("BP4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BP167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BP5:BP167").Select
    Range("BP167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BU167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BU5:BU167").Select
    Range("BU167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BZ167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BZ5:BZ167").Select
    Range("BZ167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("CE167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CE5:CE167").Select
    Range("CE167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CJ167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CJ5:CJ167").Select
    Range("CJ167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CO167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CO5:CO167").Select
    Range("CO167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CT167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CT5:CT167").Select
    Range("CT167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CY4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CY167").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CY5:CY167").Select
    Range("CY167").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CY4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CT168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CO168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CJ168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CE168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("BZ168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BU168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BP168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BK168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BF168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BA168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("AV168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AQ168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AL168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AG168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AB168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("W168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("R168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Application.CutCopyMode = False
    Selection.Cut
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("M168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("M4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("H168").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("C168").Select
    ActiveSheet.Paste
    Range("B167").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A4:B167").Select
    Range("B167").Activate
    Selection.Copy
    Range("C169").Select
    Selection.End(xlDown).Select
    Range("A3447:B3447").Select
    Range("B3447").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A168:B3447").Select
    Range("B3447").Activate
    ActiveSheet.Paste
    Range("C3447").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Columns("H:H").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 90
    Columns("H:DC").Select
    Selection.Delete Shift:=xlToLeft
    Rows("3:3").Select
    Selection.Delete Shift:=xlUp
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Target"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Respuesta"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Modalidades"
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "VARIABLE"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "INTERNET"
    Range("A2").Select
    Selection.Copy
    Range("B2").Select
    Selection.End(xlDown).Select
    Range("A3445").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A3:A3445").Select
    Range("A3445").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearFormats
    Range("C1").Select
    Selection.Copy
    Range("B1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A1").Select
    Sheets("INTERNET").Select
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Variable"
    Range("A1").Select
    Selection.End(xlToRight).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlDown).Select
    Range("A3445").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    Range("B7").Select
    ActiveWorkbook.Save
    CLEARCONTROLER(1313)
    Else
    MsgBox ERROR
End If

End Sub



'==============================================| FIN INTERNET |==========================================================

'==============================================| TELEFONIA |==========================================================

Sub TELEFONIA()



Dim NOMBRE As String
Dim ERROR As String
Dim OK As String


ERROR = "LA MACRO SELECCIONADA NO ES VALIDA"
NOMBRE = "TELEFONIA"
OK = "OK"


If ActiveSheet.Name = NOMBRE Then


'
' Macro6 Macro
'
' Acceso directo: CTRL+e
'
    CONTROLADOR(0909)
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.ColumnWidth = 40.75
    Range("B5").Select
    Selection.Cut
    Range("A6").Select
    ActiveSheet.Paste
    Range("B8").Select
    Selection.Cut
    Range("A9").Select
    ActiveSheet.Paste
    Range("B14").Select
    Selection.Cut
    Range("A15").Select
    ActiveSheet.Paste
    Range("B20").Select
    Selection.Cut
    Range("A21").Select
    ActiveSheet.Paste
    Range("B27").Select
    Selection.Cut
    Range("A28").Select
    ActiveSheet.Paste
    Range("B30").Select
    Selection.Cut
    Range("A31").Select
    ActiveSheet.Paste
    Range("B34").Select
    Selection.Cut
    Range("A35").Select
    ActiveSheet.Paste
    Range("B58").Select
    Selection.Cut
    Range("A59").Select
    ActiveSheet.Paste
    Range("B62").Select
    ActiveWindow.SmallScroll Down:=-90
    Rows("4:5").Select
    Selection.Delete Shift:=xlUp
    Rows("6:6").Select
    Selection.Delete Shift:=xlUp
    Rows("11:11").Select
    Selection.Delete Shift:=xlUp
    Rows("16:16").Select
    Selection.Delete Shift:=xlUp
    Range("B17").Select
    ActiveWindow.SmallScroll Down:=12
    Rows("22:22").Select
    Selection.Delete Shift:=xlUp
    Rows("24:24").Select
    Selection.Delete Shift:=xlUp
    Range("B20").Select
    ActiveWindow.SmallScroll Down:=9
    Rows("27:27").Select
    Selection.Delete Shift:=xlUp
    Range("A36").Select
    ActiveWindow.SmallScroll Down:=9
    Rows("50:50").Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=3
    Range("B49").Select
    ActiveWindow.SmallScroll Down:=-72
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    Columns("M:M").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("R:R").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("W:W").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    Columns("AB:AB").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AG:AG").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    Columns("AL:AL").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("AQ:AQ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 38
    Columns("AV:AV").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BA:BA").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BF:BF").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    Columns("BK:BK").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BP:BP").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    Columns("BU:BU").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("BZ:BZ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    Columns("CE:CE").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CJ:CJ").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CO:CO").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 85
    Columns("CT:CT").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("CY:CY").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 90
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 84
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 82
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("D1").Select
    Selection.Cut
    Range("C4").Select
    ActiveSheet.Paste
    Range("I1").Select
    Selection.Cut
    Range("H4").Select
    ActiveSheet.Paste
    Range("N2").Select
    Selection.Cut
    Range("M4").Select
    ActiveSheet.Paste
    Range("S2").Select
    Selection.Cut
    Range("R4").Select
    ActiveSheet.Paste
    Range("X2").Select
    Selection.Cut
    Range("W4").Select
    ActiveSheet.Paste
    Range("AC2").Select
    Selection.Cut
    Range("AB4").Select
    ActiveSheet.Paste
    Range("AH2").Select
    Selection.Cut
    Range("AG4").Select
    ActiveSheet.Paste
    Range("AM2").Select
    Selection.Cut
    Range("AL4").Select
    ActiveSheet.Paste
    Range("AR2").Select
    Selection.Cut
    Range("AQ4").Select
    ActiveSheet.Paste
    Range("AW2").Select
    Selection.Cut
    Range("AV4").Select
    ActiveSheet.Paste
    Range("BB2").Select
    Selection.Cut
    Range("BA4").Select
    ActiveSheet.Paste
    Range("BG2").Select
    Selection.Cut
    Range("BF4").Select
    ActiveSheet.Paste
    Range("BL2").Select
    Selection.Cut
    Range("BK4").Select
    ActiveSheet.Paste
    Range("BQ2").Select
    Selection.Cut
    Range("BP4").Select
    ActiveSheet.Paste
    Range("BV2").Select
    Selection.Cut
    Range("BU4").Select
    ActiveSheet.Paste
    Range("CA2").Select
    Selection.Cut
    Range("BZ4").Select
    ActiveSheet.Paste
    Range("CF2").Select
    Selection.Cut
    Range("CE4").Select
    ActiveSheet.Paste
    Range("CK2").Select
    Selection.Cut
    Range("CJ4").Select
    ActiveSheet.Paste
    Range("CP2").Select
    Selection.Cut
    Range("CO4").Select
    ActiveSheet.Paste
    Range("CU2").Select
    Selection.Cut
    Range("CT4").Select
    ActiveSheet.Paste
    Range("CZ2").Select
    Selection.Cut
    Range("CY4").Select
    ActiveSheet.Paste
    Range("C4").Select
    Selection.Copy
    Range("B4").Select
    Selection.End(xlDown).Select
    Range("C63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("C5:C63").Select
    Range("C63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("C4").Select
    Application.CutCopyMode = False
    Range("I2").Select
    Selection.Cut
    Range("H4").Select
    ActiveSheet.Paste
    Selection.Copy
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("H63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("H5:H63").Select
    Range("H63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("M63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("M5:M63").Select
    Range("M63").Activate
    ActiveSheet.Paste
    Range("M62").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("R63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("R5:R63").Select
    Range("R63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("W63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("W5:W63").Select
    Range("W63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("AB4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("AB63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AB5:AB63").Select
    Range("AB63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AG63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AG5:AG63").Select
    Range("AG63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AL63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AL5:AL63").Select
    Range("AL63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AQ63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AQ5:AQ63").Select
    Range("AQ63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("AV4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AV63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("AV5:AV63").Select
    Range("AV63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("BA63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BA5:BA63").Select
    Range("BA63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BF63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BF5:BF63").Select
    Range("BF63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("BK4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BK63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BK5:BK63").Select
    Range("BK63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BP63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BP5:BP63").Select
    Range("BP63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BU63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BU5:BU63").Select
    Range("BU63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BZ63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("BZ5:BZ63").Select
    Range("BZ63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("CE4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("CE63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CE5:CE63").Select
    Range("CE63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CJ63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CJ5:CJ63").Select
    Range("CJ63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CO63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CO5:CO63").Select
    Range("CO63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CT63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CT5:CT63").Select
    Range("CT63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CY4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CY63").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("CY5:CY63").Select
    Range("CY63").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CY4").Select
    Application.CutCopyMode = False
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CT4").Select
    Selection.End(xlDown).Select
    Range("CT64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CT4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CO4").Select
    Selection.End(xlDown).Select
    Range("CO64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CO4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CJ4").Select
    Selection.End(xlDown).Select
    Range("CJ64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Range("CJ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("CE4").Select
    Selection.End(xlDown).Select
    Range("CE64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("CE4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BZ4").Select
    Selection.End(xlDown).Select
    Range("BZ64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BZ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BU4").Select
    Selection.End(xlDown).Select
    Range("BU64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BU4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BP4").Select
    Selection.End(xlDown).Select
    Range("BP64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BP4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BK4").Select
    Selection.End(xlDown).Select
    Range("BK64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BK4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BF4").Select
    Selection.End(xlDown).Select
    Range("BF64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BF4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("BA4").Select
    Selection.End(xlDown).Select
    Range("BA64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("BA4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AV4").Select
    Selection.End(xlDown).Select
    Range("AV64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AV4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AQ4").Select
    Selection.End(xlDown).Select
    Range("AQ64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AQ4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AL4").Select
    Selection.End(xlDown).Select
    Range("AL64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AL4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AG4").Select
    Selection.End(xlDown).Select
    Range("AG64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AG4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("AB4").Select
    Selection.End(xlDown).Select
    Range("AB64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("AB4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("W4").Select
    Selection.End(xlDown).Select
    Range("W64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("W4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("R4").Select
    Selection.End(xlDown).Select
    Range("R64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("R4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("M4").Select
    Selection.End(xlDown).Select
    Range("M64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("M4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("H4").Select
    Selection.End(xlDown).Select
    Range("H64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("H4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("C64").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range("A4").Select
    Selection.Copy
    Range("A5").Select
    ActiveSheet.Paste
    Range("A6").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A7:A10").Select
    ActiveSheet.Paste
    Range("A11").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A12:A15").Select
    ActiveSheet.Paste
    Range("A16").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A17:A21").Select
    ActiveSheet.Paste
    Range("A22").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A23").Select
    ActiveSheet.Paste
    Range("A24").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A25:A26").Select
    ActiveSheet.Paste
    Range("A27").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A28:A49").Select
    ActiveSheet.Paste
    Range("A30").Select
    Selection.End(xlDown).Select
    Range("A50").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A51:A63").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("A4:B4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.End(xlUp).Select
    Range("A4").Select
    Application.CutCopyMode = False
    Range("A4:B4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("C4").Select
    Selection.End(xlDown).Select
    Range("A1263:B1263").Select
    Range("B1263").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Range("A64:B1263").Select
    Range("B1263").Activate
    ActiveSheet.Paste
    Range("B1263").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Columns("H:H").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 77
    ActiveWindow.ScrollColumn = 78
    ActiveWindow.ScrollColumn = 79
    ActiveWindow.ScrollColumn = 80
    ActiveWindow.ScrollColumn = 81
    ActiveWindow.ScrollColumn = 83
    ActiveWindow.ScrollColumn = 85
    ActiveWindow.ScrollColumn = 86
    ActiveWindow.ScrollColumn = 87
    ActiveWindow.ScrollColumn = 88
    ActiveWindow.ScrollColumn = 89
    ActiveWindow.ScrollColumn = 90
    Columns("H:DC").Select
    Selection.Delete Shift:=xlToLeft
    Rows("3:3").Select
    Selection.Delete Shift:=xlUp
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Target"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Respuesta"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Modalidades"
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Variable"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "Telefonia"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "TELEFONIA"
    Range("A2").Select
    Selection.Copy
    Range("B2").Select
    Selection.End(xlDown).Select
    Range("A1261").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A3:A1261").Select
    Range("A1261").Activate
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Range("A2:H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearFormats
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    Range("C1").Select
    Selection.Copy
    Range("B1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A9").Select
    ActiveWindow.SmallScroll Down:=-18
    Selection.End(xlDown).Select
    Range("A1261").Select
    Selection.End(xlUp).Select
    Range("H1").Select
    ActiveWorkbook.Save
    CLEARCONTROLER(1414)
    Else
    MsgBox ERROR
End If

End Sub

'==============================================| FIN TELEFONIA |==========================================================




'==============================================| EQUIPAMIENTO_MODULOS |==============================================================

Private Sub EQUIPAMIENTO_ADD_TOTAL()

    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("A5").Select
    ActiveSheet.Paste
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("B1").Select
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Columns("B:B").ColumnWidth = 13.86
    Columns("C:C").ColumnWidth = 15.29
    Columns("D:D").ColumnWidth = 12.57
    Columns("E:E").ColumnWidth = 16.57
    Range("F3:I3").Select
    Selection.Copy
    Range("B2").Select
    ActiveSheet.Paste
    Range("B1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Total"
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("B4:E4").Select
    Selection.Copy
    Range("F6").Select
    Selection.End(xlToRight).Select
    Range("CG4").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range("F4:CG4").Select
    Range("CG4").Activate
    ActiveSheet.Paste
    Range("CF4").Select
    Selection.End(xlToLeft).Select
    Range("B3").Select
    Application.CutCopyMode = False
    Range("A1").Select
End Sub


Private Sub EQUIPAMIENTO_CLEAR_TOTAL()
      Range("E31").Select
    Selection.End(xlDown).Select
    Range("A40:D40").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Rows("2:40").Select
    Range("A40").Activate
    Selection.Delete Shift:=xlUp
    Range("A40").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select

End Sub




'==============================================| FIN EQUIPAMIENTO_MODULOS |==============================================================





'==============================================| ESTILODEVIDA_MODULOS |==============================================================

 
 Private Sub ESTILODEVIDA_ADD_TOTAL()
'
' ESTILODEVIDA_ADD_TOTAL Macro
'

'
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("A5").Select
    ActiveSheet.Paste
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("B1").Select
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Columns("B:B").ColumnWidth = 16.14
    Columns("C:C").ColumnWidth = 15.57
    Columns("D:D").ColumnWidth = 20.14
    Columns("E:E").ColumnWidth = 18.57
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("F3:I3").Select
    Selection.Copy
    Range("B2").Select
    ActiveSheet.Paste
    Range("C2").Select
    Application.CutCopyMode = False
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E4").Select
    Selection.Copy
    Range("F6").Select
    Selection.End(xlToRight).Select
    Range("CG4").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range("F4:CG4").Select
    Range("CG4").Activate
    ActiveSheet.Paste
    Range("CE4").Select
    Selection.End(xlToLeft).Select
    Range("A3").Select
    Application.CutCopyMode = False
    Range("A1").Select
End Sub

Private Sub ESTILODEVIDA_CLEAR_TOTAL()
'
' ESTILODEVIDA_CLEAR_TOTAL Macro
'

'
    Range("D2").Select
    Selection.End(xlDown).Select
    Range("A345").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Rows("2:345").Select
    Range("A345").Activate
    Selection.Delete Shift:=xlUp
    Range("A343").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlToRight).Select
    Range("G1").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
End Sub



'==============================================| FIN ESTILODEVIDA_MODULOS |==============================================================


'==============================================| CONSUMOINDIVIDUO_MODULOS |==============================================================

Private Sub CONSUMOINDIVIDUO_ADD_TOTAL()

'
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("A5").Select
    ActiveSheet.Paste
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("B3").Select
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Range("D3").Select
    Columns("B:B").ColumnWidth = 19.14
    Columns("C:C").ColumnWidth = 21.86
    Columns("D:D").ColumnWidth = 29.29
    Columns("E:E").ColumnWidth = 16.29
    Range("F4").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range("F3:I3").Select
    Selection.Copy
    Range("B2").Select
    ActiveSheet.Paste
    Range("B3").Select
    Application.CutCopyMode = False
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E4").Select
    Selection.Copy
    Range("F6").Select
    Selection.End(xlToRight).Select
    Range("CG4").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range("F4:CG4").Select
    Range("CG4").Activate
    ActiveSheet.Paste
    Range("CF4").Select
    Selection.End(xlToLeft).Select
    Range("A4").Select
    Application.CutCopyMode = False
End Sub

Private Sub CONSUMOINDIVIDUO_CLEAR_TOTAL()
'
' CONSUMOINDIVIDUO_CLEAR_TOTAL Macro
'

'
    Range("E2").Select
    Selection.End(xlDown).Select
    Range("A65").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Rows("63:65").Select
    Range("A65").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Rows("2:65").Select
    Range("A65").Activate
    Selection.Delete Shift:=xlUp
    Range("A65").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
End Sub





'==============================================| FIN CONSUMOINDIVIDUO_MODULOS |==============================================================



'==============================================| CONSUMOINDIVIDUO_MARCAS_MODULOS |==============================================================
Private Sub CONSUMOINDIVIDUO_MARCA_ADD_TOTAL()

'
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("A5").Select
    ActiveSheet.Paste
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("B3").Select
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Range("D3").Select
    Columns("B:B").ColumnWidth = 19.14
    Columns("C:C").ColumnWidth = 21.86
    Columns("D:D").ColumnWidth = 29.29
    Columns("E:E").ColumnWidth = 16.29
    Range("F4").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range("F3:I3").Select
    Selection.Copy
    Range("B2").Select
    ActiveSheet.Paste
    Range("B3").Select
    Application.CutCopyMode = False
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E4").Select
    Selection.Copy
    Range("F6").Select
    Selection.End(xlToRight).Select
    Range("CG4").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range("F4:CG4").Select
    Range("CG4").Activate
    ActiveSheet.Paste
    Range("CF4").Select
    Selection.End(xlToLeft).Select
    Range("A4").Select
    Application.CutCopyMode = False


    Range("B6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("B6:E6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("F6").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A14").Select
    Selection.End(xlDown).Select
    Range("B1590:E1590").Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("B5:E1590").Select
    Range("B1590").Activate
    ActiveWindow.SmallScroll Down:=-18
    Selection.Replace What:="", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("B5").Select
    ActiveWindow.SmallScroll Down:=-24
    Range("C9").Select

    
End Sub



Private Sub CONSUMOINDIVIDUO_MARCAS_CLEAR_TOTAL()
'
' CONSUMOINDIVIDUO_MARCAS_CLEAR_TOTAL Macro
'

'
    Range("D1").Select
    ActiveWindow.SmallScroll Down:=579
    ActiveWindow.ScrollRow = 608
    ActiveWindow.ScrollRow = 675
    ActiveWindow.ScrollRow = 743
    ActiveWindow.ScrollRow = 810
    ActiveWindow.ScrollRow = 878
    ActiveWindow.ScrollRow = 1080
    ActiveWindow.ScrollRow = 1215
    ActiveWindow.ScrollRow = 1350
    ActiveWindow.ScrollRow = 1418
    ActiveWindow.ScrollRow = 1485
    ActiveWindow.ScrollRow = 1553
    ActiveWindow.ScrollRow = 1620
    ActiveWindow.ScrollRow = 1755
    ActiveWindow.ScrollRow = 1823
    ActiveWindow.ScrollRow = 2025
    ActiveWindow.ScrollRow = 2026
    ActiveWindow.SmallScroll Down:=-123
    ActiveWindow.LargeScroll Down:=-20
    ActiveWindow.SmallScroll Down:=189
    Range("A1516").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range("A2:H1516").Select
    Range("A1516").Activate
    Selection.Delete Shift:=xlUp
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
End Sub




'==============================================| FIN CONSUMOINDIVIDUO_MARCAS_MODULOS |==============================================================

'==============================================| CONSUNOHOGAR |=====================================================================
Private Sub CONSUMOHOGAR_ADD_TOTAL()

'
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("A5").Select
    ActiveSheet.Paste
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("B3").Select
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Range("D3").Select
    Columns("B:B").ColumnWidth = 19.14
    Columns("C:C").ColumnWidth = 21.86
    Columns("D:D").ColumnWidth = 29.29
    Columns("E:E").ColumnWidth = 16.29
    Range("F4").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range("F3:I3").Select
    Selection.Copy
    Range("B2").Select
    ActiveSheet.Paste
    Range("B3").Select
    Application.CutCopyMode = False
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E4").Select
    Selection.Copy
    Range("F6").Select
    Selection.End(xlToRight).Select
    Range("CG4").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range("F4:CG4").Select
    Range("CG4").Activate
    ActiveSheet.Paste
    Range("CF4").Select
    Selection.End(xlToLeft).Select
    Range("A4").Select
    Application.CutCopyMode = False
End Sub

Private Sub CONSUMOHOGAR_CLEAR_TOTAL()
'
' CONSUMOHOGAR_CLEAR_TOTAL Macro
'

'
    Range("E1").Select
    Selection.End(xlDown).Select
    Range("A89").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Rows("2:89").Select
    Range("A89").Activate
    Selection.Delete Shift:=xlUp
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Range("O1").Select
    Selection.End(xlToLeft).Select
    Range("D2").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
End Sub

Private Sub CONSUMOHOGAR_REPARAR()
'
' CONSUMOHOGAR_REPARAR Macro
'

'
    ActiveWindow.SmallScroll Down:=39
    Range("D71").Select
    ActiveWindow.SmallScroll Down:=207
    Range("D266").Select
    Selection.Copy
    Range("D266:D353").Select
    ActiveSheet.Paste
    Range("D265").Select
    Application.CutCopyMode = False
    Range("D327").Select
    Selection.End(xlUp).Select
End Sub



'==============================================| FIN CONSUNOHOGAR_MODULOS |==========================================================

'==============================================| CONSUMOHOGARMARCAS_MODULOS |==========================================================

Private Sub CONSUMOHOGARMARCAS_ADD_TOTAL()
'
' CONSUMOHOGARMARCAS_ADD_TOTAL Macro
'

'
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("A5").Select
    ActiveSheet.Paste
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("B4").Select
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Columns("B:B").ColumnWidth = 20
    Columns("C:C").ColumnWidth = 19.43
    Columns("D:D").ColumnWidth = 25
    Columns("E:E").ColumnWidth = 22.29
    Range("F3:I3").Select
    Selection.Copy
    Range("B2").Select
    ActiveSheet.Paste
    Range("C2").Select
    Application.CutCopyMode = False
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E4").Select
    Selection.Copy
    Range("F6").Select
    Selection.End(xlToRight).Select
    Range("CG4").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range("F4:CG4").Select
    Range("CG4").Activate
    ActiveSheet.Paste
    Range("CE4").Select
    Selection.End(xlToLeft).Select
    Range("B5:E5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("F6").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("A24").Select
    Selection.End(xlDown).Select
    Range("B1555:E1555").Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.Replace What:="", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("D14").Select
    Application.CutCopyMode = False
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    ActiveWindow.SmallScroll Down:=-6
    Range("B5:F5").Select
    Range("F5").Activate
    Range("F1").Select

    Range("F3").Select
    Selection.End(xlDown).Select
    Range("B5:F5").Select
    Range("F5").Activate
    Selection.ClearContents
    Range("F6").Select
    Selection.End(xlDown).Select
    Range("B23:F23").Select
    Range("F23").Activate
    Selection.ClearContents
    Range("F23").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("B34:F34").Select
    Range("F34").Activate
    Selection.ClearContents
    Range("F34").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("B51:F51").Select
    Range("F51").Activate
    Selection.ClearContents
    Range("F51").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("B63:F63").Select
    Range("F63").Activate
    Selection.ClearContents
    Range("F63").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("B79:F79").Select
    Range("F79").Activate
    Selection.ClearContents
    Range("F79").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("F94").Select
    Selection.End(xlDown).Select
    Range("B95:E95").Select
    Range("E95").Activate
    Selection.ClearContents
    Range("B113:F113").Select
    Range("F113").Activate
    Selection.ClearContents
    Range("F113").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("B124:F124").Select
    Range("F124").Activate
    Selection.ClearContents
    Range("F124").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("B137:F137").Select
    Range("F137").Activate
    Selection.ClearContents
    Range("F137").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("B144:F144").Select
    Range("F144").Activate
    Selection.ClearContents
    Range("F144").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("B156:F156").Select
    Range("F156").Activate
    Selection.ClearContents
    Range("F156").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("B169:F169").Select
    Range("F169").Activate
    Selection.ClearContents
    Range("F169").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("B187:F187").Select
    Range("F187").Activate
    Selection.ClearContents
    Range("F187").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("B195:F195").Select
    Range("F195").Activate
    Selection.ClearContents
    Range("F195").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("B211:F211").Select
    Range("F211").Activate
    Selection.ClearContents
    Range("F211").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("B225:F225").Select
    Range("F225").Activate
    Selection.ClearContents
    Range("F225").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("B247:F247").Select
    Range("F247").Activate
    Selection.ClearContents
    Range("F247").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("B266:F266").Select
    Range("F266").Activate
    Selection.ClearContents
    Range("F266").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("B287:F287").Select
    Range("F287").Activate
    Selection.ClearContents
    Range("F287").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("B306:F306").Select
    Range("F306").Activate
    Selection.ClearContents
    Range("F306").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("B323:F323").Select
    Range("F323").Activate
    Selection.Cut
    Range("B323:F323").Select
    Range("F323").Activate
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("F323").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("B339:F339").Select
    Range("F339").Activate
    Selection.ClearContents
    Range("F338").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("F3").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$CG$1555").AutoFilter Field:=6, Criteria1:="="
    ActiveWindow.SmallScroll Down:=12
    Range("B96:E112").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=57
    Range("B357:E1555").Select
    Selection.ClearContents
    ActiveWindow.SmallScroll Down:=267
    Range("E1533").Select
    Selection.AutoFilter
    ActiveWindow.SmallScroll Down:=-27
    Range("E1238").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
End Sub


Private Sub CONSUMOHOGARMARCAS_CLEAR_TOTAL()
'
' CONSUMOHOGARMARCAS_CLEAR_TOTAL Macro
'

'
    Range("D1").Select
    Selection.End(xlDown).Select
    Range("A1457").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Rows("2:1457").Select
    Range("A1457").Activate
    Selection.Delete Shift:=xlUp
    Range("A1456").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
End Sub





'==============================================| FIN CONSUMOHOGARMARCAS_MODULOS |==========================================================



'==============================================| SERVICIOSFINANCIEROS_MODULOS |==========================================================

Private Sub SERVICIOSFINANCIEROS_ADD_TOTAL()

'
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("A5").Select
    ActiveSheet.Paste
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("B3").Select
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Range("D3").Select
    Columns("B:B").ColumnWidth = 19.14
    Columns("C:C").ColumnWidth = 21.86
    Columns("D:D").ColumnWidth = 29.29
    Columns("E:E").ColumnWidth = 16.29
    Range("F4").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range("F3:I3").Select
    Selection.Copy
    Range("B2").Select
    ActiveSheet.Paste
    Range("B3").Select
    Application.CutCopyMode = False
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E4").Select
    Selection.Copy
    Range("F6").Select
    Selection.End(xlToRight).Select
    Range("CG4").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range("F4:CG4").Select
    Range("CG4").Activate
    ActiveSheet.Paste
    Range("CF4").Select
    Selection.End(xlToLeft).Select
    Range("A4").Select
    Application.CutCopyMode = False
End Sub

Private Sub SERVICIOSFINANCIEROS_CLEAR_TOTAL()
'
' SERVICIOSFINANCIEROS_CLEAR_TOTAL Macro
'

'
    Range("E2").Select
    Selection.End(xlDown).Select
    Range("A183").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Rows("2:183").Select
    Range("A183").Activate
    Selection.Delete Shift:=xlUp
    Range("A183").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select

       ActiveWindow.SmallScroll Down:=507
    Range("D547").Select
    Selection.Copy
    Range("D548").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("D548:D729").Select
    ActiveSheet.Paste
    Range("D614").Select
    Application.CutCopyMode = False
End Sub

Private Sub GENERIC_ADD_TOTAL()

'
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("A5").Select
    ActiveSheet.Paste
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("B3").Select
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Range("D3").Select
    Columns("B:B").ColumnWidth = 19.14
    Columns("C:C").ColumnWidth = 21.86
    Columns("D:D").ColumnWidth = 29.29
    Columns("E:E").ColumnWidth = 16.29
    Range("F4").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range("F3:I3").Select
    Selection.Copy
    Range("B2").Select
    ActiveSheet.Paste
    Range("B3").Select
    Application.CutCopyMode = False
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E4").Select
    Selection.Copy
    Range("F6").Select
    Selection.End(xlToRight).Select
    Range("CG4").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range("F4:CG4").Select
    Range("CG4").Activate
    ActiveSheet.Paste
    Range("CF4").Select
    Selection.End(xlToLeft).Select
    Range("A4").Select
    Application.CutCopyMode = False
End Sub

'==============================================| FIN SERVICIOSFINANCIEROS_MODULOS |==========================================================


'==============================================| TRANSPORTE_MODULOS |==========================================================
Private Sub TRANSPORTE_CLEAR_TOTAL()
'
' TRANSPORTE_CLEAR_TOTAL Macro
'

'
    Range("E2").Select
    Selection.End(xlDown).Select
    Range("A107").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Rows("2:107").Select
    Range("A107").Activate
    Selection.Delete Shift:=xlUp
    Range("A106").Select
    Selection.End(xlUp).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
End Sub


'==============================================| FIN TRANSPORTE_MODULOS |==========================================================

'==============================================| MEDIOSUP_MODULOS |==========================================================


Private Sub MEDIOSUP_CLEAR_TOTAL()
'
' MEDIOSUP_CLEAR_TOTAL Macro
'

'
    Range("E2").Select
    Selection.End(xlDown).Select
    Range("A405").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Rows("2:405").Select
    Range("A405").Activate
    Selection.Delete Shift:=xlUp
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
End Sub


'==============================================| FIN MEDIOSUP30_MODULOS |==========================================================

Private Sub MEDIOSU30_CLEAR_TOTAL()
'
' MEDIOSU30_CLEAR_TOTAL Macro
'

'
    Range("E1").Select
    Selection.End(xlDown).Select
    Range("A332").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Rows("2:332").Select
    Range("A332").Activate
    Selection.Delete Shift:=xlUp
    Range("A331").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
End Sub


'==============================================| MEDIOSDAYER_MODULOS |==========================================================
Private Sub MEDIOSDAYER_ADD_TOTAL()
'
' MEDIOSDAYER_ADD_TOTAL Macro
'

'
    Range("A3").Select
    Selection.EntireRow.Insert
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "Universo"
    Range("C3").Select
    ActiveCell.FormulaR1C1 = "Muestra"
    Range("D3").Select
    ActiveCell.FormulaR1C1 = "% Col."
    Range("E3").Select
    ActiveCell.FormulaR1C1 = "Índice"
    Range("B3:E3").Select
    Selection.AutoFill Destination:=Range("B3:BP3"), Type:=xlFillDefault
    Range("B3:BP3").Select
    Range("BN3").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range("A4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("B3").Select
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Columns("B:B").ColumnWidth = 14.71
    Columns("C:C").ColumnWidth = 12.43
    Columns("D:D").ColumnWidth = 17
    Columns("E:E").ColumnWidth = 13.86
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "Universo"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "Muestra"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "% Col."
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "Índice"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E4").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("A1").Select
End Sub

Private Sub MEDIOSDAYER_CLEAR_TOTAL()
'
' MEDIOSDAYER_CLEAR_TOTAL Macro
'

'
    Range("E10").Select
    Selection.End(xlDown).Select
    Range("A117").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Rows("116:117").Select
    Range("A117").Activate
    Range(Selection, Selection.End(xlUp)).Select
    Rows("2:117").Select
    Range("A117").Activate
    Selection.Delete Shift:=xlUp
    Range("C10").Select
End Sub

'==============================================| FIN MEDIOSDAYER_MODULOS |==========================================================

'==============================================| INTERNET_MODULOS |==========================================================
Private Sub INTERNET_CLEAR_TOTAL()
'
' INTERNET_CLEAR_TOTAL Macro
'

'
    Range("E5").Select
    Selection.End(xlDown).Select
    Range("A165").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Rows("2:165").Select
    Range("A165").Activate
    Selection.Delete Shift:=xlUp
    Range("A162").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
End Sub

'==============================================| FIN INTERNET_MODULOS |==========================================================

'==============================================| TELEFONIA_MODULOS |==================================================================
Private Sub TELEFONIA_CLEAR_TOTAL()
'
' TELEFONIA_CLEAR_TOTAL Macro
'

'
    Range("E1").Select
    Selection.End(xlDown).Select
    Range("A61").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Rows("2:61").Select
    Range("A61").Activate
    Selection.Delete Shift:=xlUp
    Range("B3").Select
End Sub

'==============================================| FIN TELEFONIA_MODULOS |==========================================================



'==================================================================================> MODULOS DE LIMPIEZA
 

Private Sub ELIMINARFORMATOCONIMG()
'ELIMINAR FORMATO DE EXCEL
    Range("A9").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Rows("1:9").Select
    Range("A9").Activate
    Selection.Delete Shift:=xlUp
    
    ActiveSheet.Shapes.Range(Array("Picture 1")).Select
    Selection.Delete
Dim rngTemp As Range
 Set rngTemp = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
 If Not rngTemp Is Nothing Then
     Range(Cells(1, 1), rngTemp).Select
  End If
  Selection.ClearFormats
End Sub


Private Sub ELIMINARFORMATONORMAL()

'QUITAR SOLO FORMATO DE CELDAS 

  Dim rngTemp As Range
  Set rngTemp = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
  If Not rngTemp Is Nothing Then
     Range(Cells(1, 1), rngTemp).Select
  End If
  Selection.ClearFormats



End Sub


Private Sub ELIMINARFORMATOCONIMGV2()
'ELIMINAR FORMATO DE EXCEL
    Range("A10").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Rows("1:10").Select
    Range("A10").Activate
    Selection.Delete Shift:=xlUp
    
    ActiveSheet.Shapes.Range(Array("Picture 1")).Select
    Selection.Delete
Dim rngTemp As Range
 Set rngTemp = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
 If Not rngTemp Is Nothing Then
     Range(Cells(1, 1), rngTemp).Select
  End If
  Selection.ClearFormats
End Sub


Private Sub ELIMINARFORMATONORMALV2()

'QUITAR SOLO FORMATO DE CELDAS 


  Dim rngTemp As Range
  Set rngTemp = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
  If Not rngTemp Is Nothing Then
     Range(Cells(1, 1), rngTemp).Select
  End If
  Selection.ClearFormats

 Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp

End Sub






'==================================================================================> FIN MODULOS DE LIMPIEZA


Private Sub CONTROLADOR(valores as Integer)

If valores = 696 then 

ELIMINARFORMATONORMALV2
'ELIMINARFORMATOCONIMGV2
MEDIOSDAYER_ADD_TOTAL

Else


'======= TIPO DE FORMATO =======
'===============================

'ELIMINARFORMATOCONIMG
     ELIMINARFORMATONORMAL
'ELIMINARFORMATOCONIMG
'VOLVER A FORMATO CON EL TOTAL
  Select Case valores
   case 0101
    PERFIL_ADD_TOTAL
   case 0202
    EQUIPAMIENTO_ADD_TOTAL
   case 0303
    ESTILODEVIDA_ADD_TOTAL
   case 0404
    CONSUMOINDIVIDUO_ADD_TOTAl
   case 0505
    CONSUMOINDIVIDUO_MARCA_ADD_TOTAL
   case 0606
    CONSUMOHOGAR_ADD_TOTAL
   case 0707
    CONSUMOHOGARMARCAS_ADD_TOTAL
   case 0808
    SERVICIOSFINANCIEROS_ADD_TOTAL
   case 0909
    GENERIC_ADD_TOTAL
   

   End Select

End if




   

End Sub



Private Sub CLEARCONTROLER(valores as Integer)



   Select Case valores

    case 0101
     PERFIL_CLEAR_TOTAL
    case 0202
     EQUIPAMIENTO_CLEAR_TOTAL
    case 0303
     ESTILODEVIDA_CLEAR_TOTAL
    case 0404
     CONSUMOINDIVIDUO_CLEAR_TOTAL
    case 0505
     CONSUMOINDIVIDUO_MARCAS_CLEAR_TOTAL
    case 0606
     CONSUMOHOGAR_CLEAR_TOTAL
     CONSUMOHOGAR_REPARAR
    case 0707
     CONSUMOHOGARMARCAS_CLEAR_TOTAL
    case 0808
     SERVICIOSFINANCIEROS_CLEAR_TOTAL
    case 0909
     TRANSPORTE_CLEAR_TOTAL
    case 1010
      MEDIOSUP_CLEAR_TOTAL
    case 1111
      MEDIOSU30_CLEAR_TOTAL
    case 1212  
     MEDIOSDAYER_CLEAR_TOTAL
    case 1313
    INTERNET_CLEAR_TOTAL
    case 1414
    TELEFONIA_CLEAR_TOTAL

   End Select

End Sub