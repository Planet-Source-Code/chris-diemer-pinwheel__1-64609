VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   6555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7185
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   437
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   479
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Randdir 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   240
   End
   Begin VB.Timer Pause 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   240
      Top             =   240
   End
   Begin VB.Label Lblinstr 
      BackColor       =   &H00000000&
      Caption         =   $"Pinwheelf.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1020
      Left            =   2760
      TabIndex        =   2
      Top             =   3030
      Width           =   3540
   End
   Begin VB.Label Lblauthor 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Chris Diemer   12/11/04"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   2760
      TabIndex        =   1
      Top             =   2520
      Width           =   3420
   End
   Begin VB.Label Lbltitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "PINWHEEL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   600
      Left            =   2880
      TabIndex        =   0
      Top             =   1770
      Width           =   3150
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   ' Press key for which part you want - KP is set to True
   ' if you press 1 - 7 or A so it won't keep restarting.
   ' Each time your press A, the program will jump to the
   ' next part until it reaches part seven and then go
   ' back to part 1.
   If KP = False And Pause.Enabled = False And Randdir.Enabled = False Then
      If KeyCode = vbKey1 Then UC = 1: Begin
      If KeyCode = vbKeyNumpad1 Then UC = 1: Begin
      If KeyCode = vbKey2 Then UC = 2: Begin
      If KeyCode = vbKeyNumpad2 Then UC = 2: Begin
      If KeyCode = vbKey3 Then UC = 3: Begin
      If KeyCode = vbKeyNumpad3 Then UC = 3: Begin
      If KeyCode = vbKey4 Then UC = 4: Begin
      If KeyCode = vbKeyNumpad4 Then UC = 4: Begin
      If KeyCode = vbKey5 Then UC = 5: Begin
      If KeyCode = vbKeyNumpad5 Then UC = 5: Begin
      If KeyCode = vbKey6 Then UC = 6: Begin
      If KeyCode = vbKeyNumpad6 Then UC = 6: Begin
      If KeyCode = vbKey7 Then UC = 7: Begin
      If KeyCode = vbKeyNumpad7 Then UC = 7: Begin
      If KeyCode = vbKeyA Then UC = 8: Begin
   End If
   If KeyCode = vbKeyEscape Then Unload Me
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   ' KP set to false after choice to be ready for next
   ' accepted input under KeyDown
   KP = False
End Sub
Private Sub Form_Load()
   ' PINWHEEL by Chris Diemer 12/11/04. I do not want this
   ' program or any part of it used in someone else's
   ' program. I have not had any problems or errors when
   ' running it, but you are free to use it at your own
   ' risk. Besides other basic languages, I have over 200
   ' programs done in Visual Basic 6. I created this about
   ' 1 1/2 years ago and I do not have a comment for all
   ' sub procedures - I could do that easily if I used
   ' comment lines at the time I wrote the program, but
   ' normally I do not have any comment lines in my programs
   ' besides the first line at the top of this paragraph.
   ShowCursor& (False) ' Shut the cursor off
   Me.Show
   Randomize ' Without this, the same randoms would be
             ' picked everytime the program was run. This
             ' uses the internal clock for a true random.
             
   ' Detect screen resolution and set borders accordingly.
   ' I like coming in 10 pixels to keep away from screen
   ' edge. The program is set to run in any screen
   ' resolution. I determine what size I want the pinwheel
   ' in 800 X 600 (which I run my desktop in) and then
   ' determine the pinwheels radius (CR) by dividing 9.23
   ' into the vertical resolution so that the pinwheel
   ' will be the same size in any resolution. More than 1
   ' variable is set by dividing into the vertical resolution
   ' because there is more than 1 screen resolution that
   ' has the same horizontal dimension and I want it
   ' proportional.
   HR = Screen.Width / Screen.TwipsPerPixelX
   VR = Screen.Height / Screen.TwipsPerPixelY
   L = 10: R = HR - 10: T = 10: B = VR - 10
   ' Variables set here for different parts of the program
   CH = HR / 2: CV = VR / 2: BG = K: DP = False
   KP = False: CR = Int(VR / 9.23)
   L1 = L + (CR * 2) + 6: R1 = R - ((CR * 2) + 6)
   T1 = T + (CR * 2) + 6: B1 = B - ((CR * 2) + 6)
   ML1 = L + CR + 3: MR1 = R - (CR + 3)
   MT1 = T + CR + 3: MB1 = B - (CR + 3)
   ML2 = L1 + CR + 3: MR2 = R1 - (CR + 3)
   MT2 = T1 + CR + 3: MB2 = B1 - (CR + 3)
   Z = Int((Rnd * 6) + S): DT = K: A = CR * 4: HN = HR - A
   VN = VR - A: RN = Int(VR / 3): RA = Int(VR / 12)
   A1 = R - L: A1 = A1 * 2: D1 = B - T: D1 = D1 * 2
   TS = A1 + D1: BI = Int(TS / 8): A1 = R - L
   D1 = Int(A1 / 3): ML = L + D1: MR = ML + D1
   RN2 = Int(VR / 6): TC = K: SM = K
   BD = Int(VR / 24): SD = BD + S
   ' Set postition of labels at runtime instead of at
   ' design time
   A1 = Int(Lbltitle.Width / 2)
   Lbltitle.Left = CH - A1
   A1 = Int(Lblauthor.Width / 2)
   Lblauthor.Left = CH - A1
   A1 = Int(Lblinstr.Width / 2)
   Lblinstr.Left = CH - A1
   Lblauthor.Top = CV - 24
   Lbltitle.Top = Lblauthor.Top - 50
   Lblinstr.Top = Lblauthor.Top + 34
   ' Set variables for introduction - movement around
   ' labels at beginning then direct path to main to run
   ' beginning part
   IL = Lblinstr.Left - 10
   IR = Lblinstr.Left + Lblinstr.Width + 10
   IT = Lbltitle.Top - 10
   IB = Lblinstr.Top + Lblinstr.Height + 10
   IL = IL - CR: IR = IR + CR: IT = IT - CR: IB = IB + CR
   Colorcount
   Lbltitle.ForeColor = RGB(O1, P1, Q1)
   Lblauthor.ForeColor = RGB(O1, P1, Q1)
   A = IL: D = IT: Setpointstwo
   O3 = O1: P3 = P1: Q3 = Q1
   DP = True
   Main
End Sub
Private Function Colorcount()
   ' This sets up the start of each new color for the full
   ' colors (O1, P1 and Q1) and the colors which will be
   ' shaded with my Lineshade color procedure (O, P and Q)
   ' The value of Z is picked at random under form load so
   ' it won't start with the same color every time. After
   ' that, it repeats the colors in order.
   Z = Z + S: If Z > S7 Then Z = S
   If Z = S And DT = K Then Z = S2
   If Z = S Then O = 40: P = 40: Q = 40: O1 = 255: P1 = 255: Q1 = 255
   If Z = S2 Then O = 40: P = 0: Q = 0: O1 = 255: P1 = 0: Q1 = 0
   If Z = S3 Then O = 0: P = 40: Q = 0: O1 = 0: P1 = 255: Q1 = 0
   If Z = S4 Then O = 0: P = 0: Q = 40: O1 = 0: P1 = 0: Q1 = 255
   If Z = S5 Then O = 40: P = 0: Q = 40: O1 = 255: P1 = 0: Q1 = 255
   If Z = S6 Then O = 40: P = 40: Q = 0: O1 = 255: P1 = 255: Q1 = 0
   If Z = S7 Then O = 40: P = 5: Q = 0: PC = 0: O1 = 255: P1 = 128: Q1 = 0
   SC = LSL
End Function
Private Function Lineshade()
   ' This is the procedure I created to shade the colors
   ' in certain parts of the program. I roll the value
   ' from 40 to 255 and back. I felt that 40 (with RGB)
   ' was as dark as I wanted to go
   If SC = LSH Then X = Y
   If SC = LSL Then X = S
   If Z = S Then
      If X = S Then O = O + S: P = P + S: Q = Q + S: SC = SC + S
      If X = Y Then O = O - S: P = P - S: Q = Q - S: SC = SC - S
   ElseIf Z = S2 Then
      If X = S Then O = O + S: SC = SC + S
      If X = Y Then O = O - S: SC = SC - S
   ElseIf Z = S3 Then
      If X = S Then P = P + S: SC = SC + S
      If X = Y Then P = P - S: SC = SC - S
   ElseIf Z = S4 Then
      If X = S Then Q = Q + S: SC = SC + S
      If X = Y Then Q = Q - S: SC = SC - S
   ElseIf Z = S5 Then
      If X = S Then O = O + S: Q = Q + S: SC = SC + S
      If X = Y Then O = O - S: Q = Q - S: SC = SC - S
   ElseIf Z = S6 Then
      If X = S Then O = O + S: P = P + S: SC = SC + S
      If X = Y Then O = O - S: P = P - S: SC = SC - S
   ElseIf Z = S7 Then
      PC = PC + S
      If X = S Then O = O + S: SC = SC + S
      If X = S And PC = Y Then P = P + S
      If X = Y Then O = O - S: SC = SC - S
      If X = Y And PC = Y And P > K Then P = P - S
      If PC = Y Then PC = K
   End If
End Function
Private Function Setpoints()
   ' Set starting points around circle for pinwheel for
   ' certain parts of the program
   CA = K: G(S) = K: G2(S) = K
   If DT = S5 Then G3(S) = K
   For C = S2 To NP
      CA = CA + CI
      G(C) = CA: G2(C) = CA
      If DT = S5 Then G3(C) = CA
   Next C
End Function
Private Function Setpointstwo()
   ' Set starting points around circle for pinwheel for
   ' certain parts of the program
   CA = K: G(S) = K
   For C = S2 To NP
      CA = CA + CI
      G(C) = CA
   Next C
End Function
Private Function Rectmoves()
   ' Move center of pinwheel around for part 1
   If A = ML1 And D < MB1 Then
      D = D + S
   ElseIf A < MR1 And D = MB1 Then
      A = A + S
   ElseIf A = MR1 And D > MT1 Then
      D = D - S
   ElseIf A > ML1 And D = MT1 Then
      A = A - S
   End If
   If E < MR2 And F = MT2 Then
      E = E + S
   ElseIf E = MR2 And F < MB2 Then
      F = F + S
   ElseIf E > ML2 And F = MB2 Then
      E = E - S
   ElseIf E = ML2 And F > MT2 Then
      F = F - S
   End If
End Function
Private Function Borderpoints()
   BC = BC + S
   If BC = BI Then
      BC = K: C = C + S
      U(C) = A1: V(C) = D1
   End If
End Function
Private Sub Begin()
   ' BG (Begin) is used to blank the labels only when the
   ' program is first started
   KP = True
   DP = False
   If BG = K Then
      BG = S
      Lbltitle.Visible = False
      Lblauthor.Visible = False
      Lblinstr.Visible = False
   End If
   Newscreen
End Sub
Private Sub Newscreen()
   ' Set variables for the current part. In part 5, I had
   ' to divide the screen into 3 parts horizontally to
   ' keep each circle in it's own area so they wouldn't
   ' end up on top of each other - this made the effect
   ' very confusing
   Cls
   RC = K: TC = K
   If UC < S8 Then DT = UC
   If UC = S8 Then
      DT = DT + S
      If DT = S8 Then DT = S
   End If
   Colorcount
   If DT = S Then
      O3 = O1: P3 = P1: Q3 = Q1: GIA = -GI
      DrawWidth = S
   ElseIf DT = S2 Then
      O3 = O: P3 = P: Q3 = Q: GIA = -GI
      DrawWidth = S2
   ElseIf DT = S3 Then
      O3 = O1: P3 = P1: Q3 = Q1: GIA = GI: MCL = S2
      DRN = S8: DrawWidth = S
   ElseIf DT = S4 Then
      O3 = O: P3 = P: Q3 = Q: GIA = GI: MCL = S2
      DRN = S8: DrawWidth = S4
   ElseIf DT = S5 Then
      O3 = O1: P3 = P1: Q3 = Q1: GIA = GI: MCL = S3
      DRN = S9: DrawWidth = S
   ElseIf DT = S6 Then
      O3 = O: P3 = P: Q3 = Q: MCL = S: BC = S: C = S
      DRN = S9: DrawWidth = S4: A1 = L: D1 = T
      U(C) = L: V(C) = T
      Do Until C = S8
         If A1 < R And D1 = T Then
            A1 = A1 + S: Borderpoints
         ElseIf A1 = R And D1 < B Then
            D1 = D1 + S: Borderpoints
         ElseIf A1 > L And D1 = B Then
            A1 = A1 - S: Borderpoints
         ElseIf A1 = L And D1 > T Then
            D1 = D1 - S: Borderpoints
         End If
      Loop
   ElseIf DT = S7 Then
      For C = S To NP: PA(C) = K: Next C
      DrawWidth = S
   End If
   If DT < S3 Then
      Line (L, T)-(R, T), RGB(O1, P1, Q1)
      Line (R, T)-(R, B), RGB(O1, P1, Q1)
      Line (R, B)-(L, B), RGB(O1, P1, Q1)
      Line (L, B)-(L, T), RGB(O1, P1, Q1)
      Line (L1, T1)-(R1, T1), RGB(O1, P1, Q1)
      Line (R1, T1)-(R1, B1), RGB(O1, P1, Q1)
      Line (R1, B1)-(L1, B1), RGB(O1, P1, Q1)
      Line (L1, B1)-(L1, T1), RGB(O1, P1, Q1)
      A = ML1: D = MT1: E = ML2: F = MT2
   ElseIf DT = S3 Or DT = S4 Then
      A = Int((Rnd * HN) + (CR * 2))
      D = Int((Rnd * VN) + (CR * 2))
      E = Int((Rnd * HN) + (CR * 2))
      F = Int((Rnd * VN) + (CR * 2))
   ElseIf DT = S5 Then
      A = L + 130
      D = Int((Rnd * VN) + (CR * 2))
      I = ML + 130
      J = Int((Rnd * VN) + (CR * 2))
      E = MR + 130
      F = Int((Rnd * VN) + (CR * 2))
   ElseIf DT = S6 Then
      A = Int((Rnd * HN) + (CR * 2))
      D = Int((Rnd * VN) + (CR * 2))
   ElseIf DT = S7 Then
      A = CH: D = CV: RC = K: ST = K
   End If
   If DT < S6 Then
      Setpoints
   ElseIf DT > S5 Then
      Setpointstwo
   End If
   If DT < S3 Or DT = S7 Then
      DP = True
      Main
   ElseIf DT = S3 Or DT = S4 Or DT = 6 Then
      Randdir.Enabled = True
   ElseIf DT = S5 Then
      Randdirtwo
   End If
End Sub
Private Sub Main()
   ' This is where the program actually runs. DP (demo in
   ' progress. DoEvents checks for other things, such as
   ' keystrokes while the program is running. With DT
   ' (draw type) or the current part, I have DT set to
   ' zero to run the introduction around the labels at the
   ' start of the program
   Do While DP = True
      DoEvents
      If Speed < GetTickCount Then
         Speed = GetTickCount + SP
         If DT = K Then
            For DSP = S To S2
               For C = S To NP
                  Line (A, D)-(A + CX(C), D + CY(C)), BackColor
                  PSet (A + CX(C), D + CY(C)), BackColor
                  G(C) = G(C) + GI
                  CX(C) = Cos(G(C)) * CR
                  CY(C) = Sin(G(C)) * CR
                  Line (A, D)-(A + CX(C), D + CY(C)), RGB(O1, P1, Q1)
               Next C
            Next DSP
            If G(S) >= GL Then Setpointstwo
            For C = S To NP
               Line (A, D)-(A + CX(C), D + CY(C)), BackColor
            Next C
            If A < IR And D = IT Then
               A = A + S
            ElseIf A = IR And D < IB Then
               D = D + S
            ElseIf A > IL And D = IB Then
               A = A - S
            ElseIf A = IL And D > IT Then
               D = D - S
            End If
            For C = S To NP
               Line (A, D)-(A + CX(C), D + CY(C)), RGB(O1, P1, Q1)
            Next C
            If A = IL And D = IT Then
               Lblinstr.ForeColor = RGB(O1, P1, Q1)
               Colorcount
               Lbltitle.ForeColor = RGB(O1, P1, Q1)
               Lblauthor.ForeColor = RGB(O1, P1, Q1)
            End If
            If SM = K Then
               SM = S: CC = 255: RC = K: E = A: F = D
               CA = K: G2(S) = K
               For C = S2 To NP
                  CA = CA + CI
                  G2(C) = CA
               Next C
            End If
            If SM = S Then
               RC = RC + S
               If RC = S2 Then
                  RC = K
                  CC = CC - S
               End If
               For C = S To NP
                  PSet (E + CX2(C), F + CY2(C)), BackColor
                  G2(C) = G2(C) + GI
                  CX2(C) = Cos(G2(C)) * CR
                  CY2(C) = Sin(G2(C)) * CR
                  PSet (E + CX2(C), F + CY2(C)), RGB(CC, CC, CC)
               Next C
            End If
            If CC = S35 Then
               SM = K
               For C = S To NP
                  PSet (E + CX2(C), F + CY2(C)), BackColor
               Next C
            End If
         ElseIf DT = S Then
            Spinlines
            For C = S To NP
               Line (A, D)-(A + CX(C), D + CY(C)), BackColor
               Line (E, F)-(E + CX2(C), F + CY2(C)), BackColor
            Next C
            Rectmoves
            For C = S To NP
               Line (A, D)-(A + CX(C), D + CY(C)), RGB(O3, P3, Q3)
               Line (E, F)-(E + CX2(C), F + CY2(C)), RGB(O3, P3, Q3)
            Next C
         ElseIf DT = S2 Then
            Spinlines
            Rectmoves
            Lineshade
            O3 = O: P3 = P: Q3 = Q
            For C = S To NP
               Line (A, D)-(A + CX(C), D + CY(C)), RGB(O3, P3, Q3)
               Line (E, F)-(E + CX2(C), F + CY2(C)), RGB(O3, P3, Q3)
            Next C
         ElseIf DT = S3 Then
            Cls
            Spinlines
            DC = DC + S
            A = A + AA: D = D + DA
            E = E + EA: F = F + FA
         ElseIf DT = S4 Then
            Lineshade
            O3 = O: P3 = P: Q3 = Q
            Spinlines
            DC = DC + S
            A = A + AA: D = D + DA
            E = E + EA: F = F + FA
         ElseIf DT = S5 Then
            Cls
            Spinlines
            DC = DC + S
            A = A + AA: D = D + DA
            E = E + EA: F = F + FA
            I = I + IA: J = J + JA
         ElseIf DT = S6 Then
            Lineshade
            O3 = O: P3 = P: Q3 = Q
            Spinlinestwo
            DC = DC + S
            A = A + AA: D = D + DA
         ElseIf DT = S7 Then
            For DSP = S To S2
               For C = S To NP
                  Line (A, D)-(A + CX(C), D + CY(C)), BackColor
                  G(C) = G(C) + GI
                  CX(C) = Cos(G(C)) * CR
                  CY(C) = Sin(G(C)) * CR
                  Line (A, D)-(A + CX(C), D + CY(C)), RGB(O1, P1, Q1)
               Next C
            Next DSP
            For C = S To NP
               If PA(C) = K Then
                  If A + CX(C) <= A And D + CY(C) <= D Then
                     M(C) = A + CX(C): N(C) = D + CY(C)
                     DR = Int((Rnd * S3) + S)
                     If DR = S Then MA(C) = -S: NA(C) = S
                     If DR = S2 Then MA(C) = -S: NA(C) = K
                     If DR = S3 Then MA(C) = -S: NA(C) = -S
                  ElseIf A + CX(C) > A And D + CY(C) <= D Then
                     M(C) = A + CX(C): N(C) = D + CY(C)
                     DR = Int((Rnd * S3) + S)
                     If DR = S Then MA(C) = S: NA(C) = -S
                     If DR = S2 Then MA(C) = S: NA(C) = K
                     If DR = S3 Then MA(C) = S: NA(C) = S
                  ElseIf A + CX(C) <= A And D + CY(C) > D Then
                     M(C) = A + CX(C): N(C) = D + CY(C)
                     DR = Int((Rnd * S3) + S)
                     If DR = S Then MA(C) = -S: NA(C) = S
                     If DR = S2 Then MA(C) = -S: NA(C) = K
                     If DR = S3 Then MA(C) = -S: NA(C) = -S
                  ElseIf A + CX(C) > A And D + CY(C) > D Then
                     M(C) = A + CX(C): N(C) = D + CY(C)
                     DR = Int((Rnd * S3) + S)
                     If DR = S Then MA(C) = S: NA(C) = -S
                     If DR = S2 Then MA(C) = S: NA(C) = K
                     If DR = S3 Then MA(C) = S: NA(C) = S
                  End If
                  PA(C) = S
               End If
            Next C
            For C = S To NP
               If PA(C) = S Then
                  PSet (M(C), N(C)), BackColor
                  M(C) = M(C) + MA(C)
                  N(C) = N(C) + NA(C)
                  PSet (M(C), N(C)), vbWhite
               End If
               If M(C) = L + BD Or M(C) = R - BD Or N(C) = T + BD Or N(C) = B - BD Then
                  If C = S And PA(C) = S Then
                     For H = S To NP
                        M1(H) = M(C): N1(H) = N(C)
                     Next H
                     PA(C) = S2: SS(C) = K
                  ElseIf C = S2 And PA(C) = S Then
                     For H = S To NP
                        M2(H) = M(C): N2(H) = N(C)
                     Next H
                     PA(C) = S2: SS(C) = K
                  ElseIf C = S3 And PA(C) = S Then
                     For H = S To NP
                        M3(H) = M(C): N3(H) = N(C)
                     Next H
                     PA(C) = S2: SS(C) = K
                  ElseIf C = S4 And PA(C) = S Then
                     For H = S To NP
                        M4(H) = M(C): N4(H) = N(C)
                     Next H
                     PA(C) = S2: SS(C) = K
                  ElseIf C = S5 And PA(C) = S Then
                     For H = S To NP
                        M5(H) = M(C): N5(H) = N(C)
                     Next H
                     PA(C) = S2: SS(C) = K
                  ElseIf C = S6 And PA(C) = S Then
                     For H = S To NP
                        M6(H) = M(C): N6(H) = N(C)
                     Next H
                     PA(C) = S2: SS(C) = K
                  ElseIf C = S7 And PA(C) = S Then
                     For H = S To NP
                        M7(H) = M(C): N7(H) = N(C)
                     Next H
                     PA(C) = S2: SS(C) = K
                  ElseIf C = S8 And PA(C) = S Then
                     For H = S To NP
                        M8(H) = M(C): N8(H) = N(C)
                     Next H
                     PA(C) = S2: SS(C) = K
                  End If
               End If
               If PA(C) = S2 Then
                  If C = S Then
                     SS(C) = SS(C) + S
                     If SS(C) > K And SS(C) < SD Then
                        For H = S To NP
                           PSet (M1(H), N1(H)), BackColor
                        Next H
                        M1(S) = M1(S) - S: M1(S2) = M1(S2) + S
                        N1(S3) = N1(S3) - S: N1(S4) = N1(S4) + S
                        M1(S5) = M1(S5) - S: N1(S5) = N1(S5) - S
                        M1(S6) = M1(S6) + S: N1(S6) = N1(S6) - S
                        M1(S7) = M1(S7) - S: N1(S7) = N1(S7) + S
                        M1(S8) = M1(S8) + S: N1(S8) = N1(S8) + S
                        For H = S To NP
                           PSet (M1(H), N1(H)), vbWhite
                        Next H
                     ElseIf SS(C) = SD Then
                        For H = S To NP
                           PSet (M1(H), N1(H)), BackColor
                        Next H
                        PA(C) = S3
                     End If
                  ElseIf C = S2 Then
                     SS(C) = SS(C) + S
                     If SS(C) > K And SS(C) < SD Then
                        For H = S To NP
                           PSet (M2(H), N2(H)), BackColor
                        Next H
                        M2(S) = M2(S) - S: M2(S2) = M2(S2) + S
                        N2(S3) = N2(S3) - S: N2(S4) = N2(S4) + S
                        M2(S5) = M2(S5) - S: N2(S5) = N2(S5) - S
                        M2(S6) = M2(S6) + S: N2(S6) = N2(S6) - S
                        M2(S7) = M2(S7) - S: N2(S7) = N2(S7) + S
                        M2(S8) = M2(S8) + S: N2(S8) = N2(S8) + S
                        For H = S To NP
                           PSet (M2(H), N2(H)), vbWhite
                        Next H
                     ElseIf SS(C) = SD Then
                        For H = S To NP
                           PSet (M2(H), N2(H)), BackColor
                        Next H
                        PA(C) = S3
                     End If
                  ElseIf C = S3 Then
                     SS(C) = SS(C) + S
                     If SS(C) > K And SS(C) < SD Then
                        For H = S To NP
                           PSet (M3(H), N3(H)), BackColor
                        Next H
                        M3(S) = M3(S) - S: M3(S2) = M3(S2) + S
                        N3(S3) = N3(S3) - S: N3(S4) = N3(S4) + S
                        M3(S5) = M3(S5) - S: N3(S5) = N3(S5) - S
                        M3(S6) = M3(S6) + S: N3(S6) = N3(S6) - S
                        M3(S7) = M3(S7) - S: N3(S7) = N3(S7) + S
                        M3(S8) = M3(S8) + S: N3(S8) = N3(S8) + S
                        For H = S To NP
                           PSet (M3(H), N3(H)), vbWhite
                        Next H
                     ElseIf SS(C) = SD Then
                        For H = S To NP
                           PSet (M3(H), N3(H)), BackColor
                        Next H
                        PA(C) = S3
                     End If
                  ElseIf C = S4 Then
                     SS(C) = SS(C) + S
                     If SS(C) > K And SS(C) < SD Then
                        For H = S To NP
                           PSet (M4(H), N4(H)), BackColor
                        Next H
                        M4(S) = M4(S) - S: M4(S2) = M4(S2) + S
                        N4(S3) = N4(S3) - S: N4(S4) = N4(S4) + S
                        M4(S5) = M4(S5) - S: N4(S5) = N4(S5) - S
                        M4(S6) = M4(S6) + S: N4(S6) = N4(S6) - S
                        M4(S7) = M4(S7) - S: N4(S7) = N4(S7) + S
                        M4(S8) = M4(S8) + S: N4(S8) = N4(S8) + S
                        For H = S To NP
                           PSet (M4(H), N4(H)), vbWhite
                        Next H
                     ElseIf SS(C) = SD Then
                        For H = S To NP
                           PSet (M4(H), N4(H)), BackColor
                        Next H
                        PA(C) = S3
                     End If
                  ElseIf C = S5 Then
                     SS(C) = SS(C) + S
                     If SS(C) > K And SS(C) < SD Then
                        For H = S To NP
                           PSet (M5(H), N5(H)), BackColor
                        Next H
                        M5(S) = M5(S) - S: M5(S2) = M5(S2) + S
                        N5(S3) = N5(S3) - S: N5(S4) = N5(S4) + S
                        M5(S5) = M5(S5) - S: N5(S5) = N5(S5) - S
                        M5(S6) = M5(S6) + S: N5(S6) = N5(S6) - S
                        M5(S7) = M5(S7) - S: N5(S7) = N5(S7) + S
                        M5(S8) = M5(S8) + S: N5(S8) = N5(S8) + S
                        For H = S To NP
                           PSet (M5(H), N5(H)), vbWhite
                        Next H
                     ElseIf SS(C) = SD Then
                        For H = S To NP
                           PSet (M5(H), N5(H)), BackColor
                        Next H
                        PA(C) = S3
                     End If
                  ElseIf C = S6 Then
                     SS(C) = SS(C) + S
                     If SS(C) > K And SS(C) < SD Then
                        For H = S To NP
                           PSet (M6(H), N6(H)), BackColor
                        Next H
                        M6(S) = M6(S) - S: M6(S2) = M6(S2) + S
                        N6(S3) = N6(S3) - S: N6(S4) = N6(S4) + S
                        M6(S5) = M6(S5) - S: N6(S5) = N6(S5) - S
                        M6(S6) = M6(S6) + S: N6(S6) = N6(S6) - S
                        M6(S7) = M6(S7) - S: N6(S7) = N6(S7) + S
                        M6(S8) = M6(S8) + S: N6(S8) = N6(S8) + S
                        For H = S To NP
                           PSet (M6(H), N6(H)), vbWhite
                        Next H
                     ElseIf SS(C) = SD Then
                        For H = S To NP
                           PSet (M6(H), N6(H)), BackColor
                        Next H
                        PA(C) = S3
                     End If
                  ElseIf C = S7 Then
                     SS(C) = SS(C) + S
                     If SS(C) > K And SS(C) < SD Then
                        For H = S To NP
                           PSet (M7(H), N7(H)), BackColor
                        Next H
                        M7(S) = M7(S) - S: M7(S2) = M7(S2) + S
                        N7(S3) = N7(S3) - S: N7(S4) = N7(S4) + S
                        M7(S5) = M7(S5) - S: N7(S5) = N7(S5) - S
                        M7(S6) = M7(S6) + S: N7(S6) = N7(S6) - S
                        M7(S7) = M7(S7) - S: N7(S7) = N7(S7) + S
                        M7(S8) = M7(S8) + S: N7(S8) = N7(S8) + S
                        For H = S To NP
                           PSet (M7(H), N7(H)), vbWhite
                        Next H
                     ElseIf SS(C) = SD Then
                        For H = S To NP
                           PSet (M7(H), N7(H)), BackColor
                        Next H
                        PA(C) = S3
                     End If
                  ElseIf C = S8 Then
                     SS(C) = SS(C) + S
                     If SS(C) > K And SS(C) < SD Then
                        For H = S To NP
                           PSet (M8(H), N8(H)), BackColor
                        Next H
                        M8(S) = M8(S) - S: M8(S2) = M8(S2) + S
                        N8(S3) = N8(S3) - S: N8(S4) = N8(S4) + S
                        M8(S5) = M8(S5) - S: N8(S5) = N8(S5) - S
                        M8(S6) = M8(S6) + S: N8(S6) = N8(S6) - S
                        M8(S7) = M8(S7) - S: N8(S7) = N8(S7) + S
                        M8(S8) = M8(S8) + S: N8(S8) = N8(S8) + S
                        For H = S To NP
                           PSet (M8(H), N8(H)), vbWhite
                        Next H
                     ElseIf SS(C) = SD Then
                        For H = S To NP
                           PSet (M8(H), N8(H)), BackColor
                        Next H
                        PA(C) = S3
                     End If
                  End If
               End If
               If PA(C) = S3 Then
                  PA(C) = K
                  RC = RC + S
                  If RC = S72 Then
                     RC = K
                     ST = ST + S
                     If ST = S Then
                        Colorcount
                     ElseIf ST = S2 Then
                        DP = False
                        Pause.Enabled = True
                     End If
                  End If
               End If
            Next C
         End If
         If G(S) >= GL Then Setpoints
         If DT < S3 And A = ML1 And D = MT1 Then
            DP = False
            Pause.Enabled = True
         End If
         If DT > S2 And DT < S7 And DC = DS Then
            RC = RC + S
            If RC < LSL Then
               If DT = S3 Or DT = S4 Or DT = S6 Then
                  DP = False
                  Randdir.Enabled = True
               ElseIf DT = S5 Then
                  DP = False
                  Randdirtwo
               End If
            ElseIf RC = LSL Then
               DP = False
               Pause.Enabled = True
            End If
         End If
      End If
   Loop
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   ' I always use QueryUnload when running a loop like in
   ' Main with DoEvents. Turn the cursor back on when the
   ' program ends
   ShowCursor& (True)
   End
End Sub
Private Sub Pause_Timer()
   ' I like a short pause between parts here, instead of
   ' jumping instantly to the next part or the same part
   ' (with a different color) if that is what you have
   ' chosen
   TC = TC + S
   If TC = S Then
      Pause.Interval = 500
      Cls
   ElseIf TC = S2 Then
      Pause.Interval = 2000
      Pause.Enabled = False
      Newscreen
   End If
End Sub
Private Sub Randdir_Timer()
   ' This makes sure I can move a random number of spaces
   ' in a random direction without going out of bounds
   ' ( off the screen)
   DC = K: MC = K
   DS = Int((Rnd * RN) + RA)
   DR = Int((Rnd * DRN) + S)
   If DR = S And (A - DS) > (L + CR) Then
      AA = -S: DA = K: MC = MC + S
   ElseIf DR = S2 And (A + DS) < (R - CR) Then
      AA = S: DA = K: MC = MC + S
   ElseIf DR = S3 And (D - DS) > (T + CR) Then
      AA = K: DA = -S: MC = MC + S
   ElseIf DR = S4 And (D + DS) < (B - CR) Then
      AA = K: DA = S: MC = MC + S
   ElseIf DR = S5 And (A - DS) > (L + CR) And (D - DS) > (T + CR) Then
      AA = -S: DA = -S: MC = MC + S
   ElseIf DR = S6 And (A + DS) < (R - CR) And (D - DS) > (T + CR) Then
      AA = S: DA = -S: MC = MC + S
   ElseIf DR = S7 And (A - DS) > (L + CR) And (D + DS) < (B - CR) Then
      AA = -S: DA = S: MC = MC + S
   ElseIf DR = S8 And (A + DS) < (R - CR) And (D + DS) < (B - CR) Then
      AA = S: DA = S: MC = MC + S
   ElseIf DR = S9 Then
      AA = K: DA = K: MC = MC + S
   End If
   If DT = S3 Or DT = S4 Then
      DR = Int((Rnd * DRN) + S)
      If DR = S And (E - DS) > (L + CR) Then
         EA = -S: FA = K: MC = MC + S
      ElseIf DR = S2 And (E + DS) < (R - CR) Then
         EA = S: FA = K: MC = MC + S
      ElseIf DR = S3 And (F - DS) > (T + CR) Then
         EA = K: FA = -S: MC = MC + S
      ElseIf DR = S4 And (F + DS) < (B - CR) Then
         EA = K: FA = S: MC = MC + S
      ElseIf DR = S5 And (E - DS) > (L + CR) And (F - DS) > (T + CR) Then
         EA = -S: FA = -S: MC = MC + S
      ElseIf DR = S6 And (E + DS) < (R - CR) And (F - DS) > (T + CR) Then
         EA = S: FA = -S: MC = MC + S
      ElseIf DR = S7 And (E - DS) > (L + CR) And (F + DS) < (B - CR) Then
         EA = -S: FA = S: MC = MC + S
      ElseIf DR = S8 And (E + DS) < (R - CR) And (F + DS) < (B - CR) Then
         EA = S: FA = S: MC = MC + S
      ElseIf DR = S9 Then
         EA = K: FA = K: MC = MC + S
      End If
   End If
   ' MC lets me know if the move is clear
   If MC = MCL Then
      If DT = S6 And DR = S9 Then
         TS = Int(DS * 1.5)
         DS = TS
      End If
      Randdir.Enabled = False
      DP = True
      Main
   End If
End Sub
Private Sub Spinlines()
   ' Here I am using DSP with a for/next to make the lines
   ' spin twice as fast as the center of the circle they
   ' are rotating around
   For DSP = S To S2
      For C = S To NP
         If DT = S Then
            Line (A, D)-(A + CX(C), D + CY(C)), BackColor
            Line (E, F)-(E + CX2(C), F + CY2(C)), BackColor
         End If
         If DT = S5 Then
            Line (I, J)-(I + CX3(C), J + CY3(C)), BackColor
            G3(C) = G3(C) + GI
            CX3(C) = Cos(G3(C)) * CR
            CY3(C) = Sin(G3(C)) * CR
         End If
         G(C) = G(C) + GI
         CX(C) = Cos(G(C)) * CR
         CY(C) = Sin(G(C)) * CR
         G2(C) = G2(C) + GIA
         CX2(C) = Cos(G2(C)) * CR
         CY2(C) = Sin(G2(C)) * CR
         Line (A, D)-(A + CX(C), D + CY(C)), RGB(O3, P3, Q3)
         Line (E, F)-(E + CX2(C), F + CY2(C)), RGB(O3, P3, Q3)
         If DT = S5 Then
            Line (I, J)-(I + CX3(C), J + CY3(C)), RGB(O3, P3, Q3)
         End If
         If DT > S2 And DT < S5 And DSP = S2 Then
            Line (A + CX(C), D + CY(C))-(E + CX2(C), F + CY2(C)), RGB(O3, P3, Q3)
         End If
         If DT = S5 And DSP = S2 Then
            Line (A + CX(C), D + CY(C))-(I + CX3(C), J + CY3(C)), RGB(O3, P3, Q3)
            Line (I + CX3(C), J + CY3(C))-(E + CX2(C), F + CY2(C)), RGB(O3, P3, Q3)
         End If
      Next C
   Next DSP
End Sub
Private Sub Spinlinestwo()
   ' This is similar to the above prodecure - each is used
   ' for a different part of the program
   For DSP = S To S2
      For C = S To NP
         G(C) = G(C) + GI
         CX(C) = Cos(G(C)) * CR
         CY(C) = Sin(G(C)) * CR
         Line (A, D)-(A + CX(C), D + CY(C)), RGB(O3, P3, Q3)
      Next C
   Next DSP
   For C = S To NP
      If U(C) < R And V(C) = T Then
         U(C) = U(C) + S
      ElseIf U(C) = R And V(C) < B Then
         V(C) = V(C) + S
      ElseIf U(C) > L And V(C) = B Then
         U(C) = U(C) - S
      ElseIf U(C) = L And V(C) > T Then
         V(C) = V(C) - S
      End If
      Line (U(C), V(C))-(A + CX(C), D + CY(C)), RGB(O3, P3, Q3)
   Next C
End Sub
Private Sub Randdirtwo()
   ' This makes sure I can move a random number of spaces
   ' in a random direction without going out of bounds
   ' ( off the screen)
   DC = K: MC = K
   DS = Int((Rnd * RN2) + RA)
   DR = Int((Rnd * DRN) + S)
   If DR = S And (A - DS) > (L + CR) Then
      AA = -S: DA = K: MC = MC + S
   ElseIf DR = S2 And (A + DS) < (ML - CR) Then
      AA = S: DA = K: MC = MC + S
   ElseIf DR = S3 And (D - DS) > (T + CR) Then
      AA = K: DA = -S: MC = MC + S
   ElseIf DR = S4 And (D + DS) < (B - CR) Then
      AA = K: DA = S: MC = MC + S
   ElseIf DR = S5 And (A - DS) > (L + CR) And (D - DS) > (T + CR) Then
      AA = -S: DA = -S: MC = MC + S
   ElseIf DR = S6 And (A + DS) < (ML - CR) And (D - DS) > (T + CR) Then
      AA = S: DA = -S: MC = MC + S
   ElseIf DR = S7 And (A - DS) > (L + CR) And (D + DS) < (B - CR) Then
      AA = -S: DA = S: MC = MC + S
   ElseIf DR = S8 And (A + DS) < (ML - CR) And (D + DS) < (B - CR) Then
      AA = S: DA = S: MC = MC + S
   ElseIf DR = S9 Then
      AA = K: DA = K: MC = MC + S
   End If
   Randdirtwocheck
End Sub
Private Sub Randdirtwocheck()
   If MC = S Then
      Randdirthree
   ElseIf MC = K Then
      Randdirtwo
   End If
End Sub
Private Sub Randdirthree()
   ' This makes sure I can move a random number of spaces
   ' in a random direction without going out of bounds
   ' ( off the screen)
   DR = Int((Rnd * DRN) + S)
   If DR = S And (E - DS) > (MR + CR) Then
      EA = -S: FA = K: MC = MC + S
   ElseIf DR = S2 And (E + DS) < (R - CR) Then
      EA = S: FA = K: MC = MC + S
   ElseIf DR = S3 And (F - DS) > (T + CR) Then
      EA = K: FA = -S: MC = MC + S
   ElseIf DR = S4 And (F + DS) < (B - CR) Then
      EA = K: FA = S: MC = MC + S
   ElseIf DR = S5 And (E - DS) > (MR + CR) And (F - DS) > (T + CR) Then
      EA = -S: FA = -S: MC = MC + S
   ElseIf DR = S6 And (E + DS) < (R - CR) And (F - DS) > (T + CR) Then
      EA = S: FA = -S: MC = MC + S
   ElseIf DR = S7 And (E - DS) > (MR + CR) And (F + DS) < (B - CR) Then
      EA = -S: FA = S: MC = MC + S
   ElseIf DR = S8 And (E + DS) < (R - CR) And (F + DS) < (B - CR) Then
      EA = S: FA = S: MC = MC + S
   ElseIf DR = S9 Then
      EA = K: FA = K: MC = MC + S
   End If
   Randdirthreecheck
End Sub
Private Sub Randdirthreecheck()
   If MC = S2 Then
      Randdirfour
   ElseIf MC = S Then
      Randdirthree
   End If
End Sub
Private Sub Randdirfour()
   ' This makes sure I can move a random number of spaces
   ' in a random direction without going out of bounds
   ' ( off the screen)
   DR = Int((Rnd * DRN) + S)
   If DR = S And (I - DS) > (ML + CR) Then
      IA = -S: JA = K: MC = MC + S
   ElseIf DR = S2 And (I + DS) < (MR - CR) Then
      IA = S: JA = K: MC = MC + S
   ElseIf DR = S3 And (J - DS) > (T + CR) Then
      IA = K: JA = -S: MC = MC + S
   ElseIf DR = S4 And (J + DS) < (B - CR) Then
      IA = K: JA = S: MC = MC + S
   ElseIf DR = S5 And (I - DS) > (ML + CR) And (J - DS) > (T + CR) Then
      IA = -S: JA = -S: MC = MC + S
   ElseIf DR = S6 And (I + DS) < (MR - CR) And (J - DS) > (T + CR) Then
      IA = S: JA = -S: MC = MC + S
   ElseIf DR = S7 And (I - DS) > (ML + CR) And (J + DS) < (B - CR) Then
      IA = -S: JA = S: MC = MC + S
   ElseIf DR = S8 And (I + DS) < (MR - CR) And (J + DS) < (B - CR) Then
      IA = S: JA = S: MC = MC + S
   ElseIf DR = S9 Then
      IA = K: JA = K: MC = MC + S
   End If
   Randdirfourcheck
End Sub
Private Sub Randdirfourcheck()
   If MC = MCL Then
      DP = True
      Main
   ElseIf MC = S2 Then
      Randdirfour
   End If
End Sub
