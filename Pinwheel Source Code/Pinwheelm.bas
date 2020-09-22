Attribute VB_Name = "Module1"
Declare Function ShowCursor& Lib "user32" (ByVal bShow As Long)
Declare Function GetTickCount Lib "kernel32" () As Long
Public HR As Integer, VR As Integer, CH As Integer, CV As Integer, L As Integer, R As Integer
Public T As Integer, B As Integer, BG As Byte, DP As Boolean, CR As Single, A As Single
Public D As Single, CA As Single, DT As Byte, UC As Byte, KP As Boolean, L1 As Integer
Public R1 As Integer, T1 As Integer, B1 As Integer, ML1 As Integer, MR1 As Integer, MT1 As Integer
Public MB1 As Integer, ML2 As Integer, MR2 As Integer, MT2 As Integer, MB2 As Integer
Public E As Single, F As Single, C As Byte, Z As Byte, SC As Byte, PC As Byte, O As Integer
Public P As Integer, Q As Integer, O1 As Integer, P1 As Integer, Q1 As Integer, X As Byte
Public Speed As Long, DSP As Byte, RC As Byte, O3 As Integer, P3 As Integer, Q3 As Integer
Public TC As Integer, DC As Integer, DS As Integer, DR As Byte, MC As Byte, HN As Integer
Public VN As Integer, RN As Integer, RA As Integer, AA As Integer, DA As Integer, EA As Integer
Public FA As Integer, GIA As Single, I As Single, J As Single, IA As Integer, JA As Integer
Public MCL As Byte, DRN As Byte, A1 As Integer, D1 As Integer, TS As Integer, BI As Integer
Public BC As Integer, ML As Integer, MR As Integer, RN2 As Integer, IL As Integer, IR As Integer
Public IT As Integer, IB As Integer, SM As Byte, CC As Integer, ST As Byte, H As Byte
Public BD As Integer, SD As Integer
Public Const K As Byte = 0, S As Byte = 1, SP As Byte = 1, GI As Single = 0.01
Public Const GL As Single = 6.29, CI As Single = 0.78625, NP As Byte = 8, S2 As Byte = 2
Public Const S3 As Byte = 3, S4 As Byte = 4, S5 As Byte = 5, S6 As Byte = 6, S7 As Byte = 7
Public Const S8 As Byte = 8, S9 As Byte = 9, S35 As Byte = 35, Y As Byte = 2, LSL As Byte = 40
Public Const LSH As Byte = 255, S72 As Byte = 72
Public CX(1 To 8) As Single, CY(1 To 8) As Single, CX2(1 To 8) As Single, CY2(1 To 8) As Single
Public CX3(1 To 8) As Single, CY3(1 To 8) As Single, PA(1 To 8) As Byte
Public G(1 To 8) As Single, G2(1 To 8) As Single, G3(1 To 8) As Single
Public U(1 To 8) As Integer, V(1 To 8) As Integer, M(1 To 8) As Integer, N(1 To 8) As Integer
Public MA(1 To 8) As Integer, NA(1 To 8) As Integer, SS(1 To 8) As Byte
Public M1(1 To 8) As Integer, N1(1 To 8) As Integer, M2(1 To 8) As Integer, N2(1 To 8) As Integer
Public M3(1 To 8) As Integer, N3(1 To 8) As Integer, M4(1 To 8) As Integer, N4(1 To 8) As Integer
Public M5(1 To 8) As Integer, N5(1 To 8) As Integer, M6(1 To 8) As Integer, N6(1 To 8) As Integer
Public M7(1 To 8) As Integer, N7(1 To 8) As Integer, M8(1 To 8) As Integer, N8(1 To 8) As Integer




















































































