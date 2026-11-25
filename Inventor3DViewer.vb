Imports System.Windows.Forms

Imports DrawingColor = System.Drawing.Color
Imports DrawingPoint = System.Drawing.Point
Imports DrawingSize = System.Drawing.Size
Imports DrawingFont = System.Drawing.Font
Imports DrawingFontStyle = System.Drawing.FontStyle

Imports IOFile = System.IO.File
Imports IOPath = System.IO.Path


Public NotInheritable Class Inventor3DViewer


    Private Sub New()

    End Sub


    '================================================================
    ' PUBLIC FUNCTION
    '
    ' CALL THIS FUNCTION FROM ANY FORM
    '================================================================

    Public Shared Sub ShowModel(
        owner As IWin32Window,
        modelPath As String
    )

        Try

            '========================================================
            ' CHECK MODEL PATH
            '========================================================

            If String.IsNullOrWhiteSpace(modelPath) Then

                MsgBox(
                    "Model path is empty."
                )

                Exit Sub

            End If


            If Not IOFile.Exists(modelPath) Then

                MsgBox(
                    "3D model was not found:" &
                    vbCrLf &
                    vbCrLf &
                    modelPath
                )

                Exit Sub

            End If


            '========================================================
            ' OPEN VIEWER POPUP
            '========================================================

            Using viewerForm As New ModelViewerForm(
                modelPath
            )

                viewerForm.ShowDialog(
                    owner
                )

            End Using


        Catch ex As Exception

            MsgBox(
                "Unable to open 3D model viewer:" &
                vbCrLf &
                vbCrLf &
                ex.ToString()
            )

        End Try

    End Sub



    '================================================================
    ' CUSTOM FOCUSABLE PANEL
    '
    ' THIS ALLOWS MOUSE WHEEL EVENTS
    '================================================================

    Private Class ViewerSurface
        Inherits Panel


        Public Sub New()

            Me.SetStyle(
                ControlStyles.Selectable,
                True
            )

            Me.TabStop =
                True

            Me.BackColor =
                DrawingColor.White

        End Sub


        Protected Overrides Sub OnMouseDown(
            e As MouseEventArgs
        )

            Me.Focus()

            MyBase.OnMouseDown(
                e
            )

        End Sub

    End Class



    '================================================================
    ' VIEWER POPUP FORM
    '================================================================

    Private Class ModelViewerForm
        Inherits Form


        '============================================================
        ' MODEL PATH
        '============================================================

        Private ReadOnly _modelPath As String


        '============================================================
        ' APPRENTICE OBJECTS
        '============================================================

        Private apprentice As Inventor.ApprenticeServerComponent =
            Nothing

        Private apprenticeDocument As Inventor.ApprenticeServerDocument =
            Nothing

        Private clientView As Inventor.ClientView =
            Nothing


        '============================================================
        ' UI CONTROLS
        '============================================================

        Private headerPanel As Panel

        Private bottomPanel As Panel

        Private modelPanel As ViewerSurface


        Private titleLabel As Label

        Private modelNameLabel As Label

        Private helpLabel As Label


        Private frontButton As Button

        Private topButton As Button

        Private isoButton As Button

        Private fitButton As Button

        Private closeButton As Button


        '============================================================
        ' MOUSE CONTROL
        '============================================================

        Private mouseDragging As Boolean =
            False

        Private currentMouseOperation As String =
            ""

        Private previousMousePoint As DrawingPoint



        '============================================================
        ' CONSTRUCTOR
        '============================================================

        Public Sub New(
            modelPath As String
        )

            _modelPath =
                modelPath

            InitializeViewerUI()

        End Sub



        '============================================================
        ' CREATE VIEWER UI
        '============================================================

        Private Sub InitializeViewerUI()


            '========================================================
            ' FORM
            '========================================================

            Me.Text =
                "3D Model Viewer"


            Me.StartPosition =
                FormStartPosition.CenterParent


            Me.Size =
                New DrawingSize(
                    900,
                    650
                )


            Me.MinimumSize =
                New DrawingSize(
                    750,
                    520
                )


            Me.FormBorderStyle =
                FormBorderStyle.Sizable


            Me.MaximizeBox =
                False


            Me.MinimizeBox =
                False


            Me.ShowIcon =
                False


            Me.BackColor =
                DrawingColor.White


            Me.Font =
                New DrawingFont(
                    "Segoe UI",
                    9.0F,
                    DrawingFontStyle.Regular
                )



            '========================================================
            ' HEADER PANEL
            '========================================================

            headerPanel =
                New Panel()


            headerPanel.Dock =
                DockStyle.Top


            headerPanel.Height =
                60


            headerPanel.BackColor =
                DrawingColor.White


            Me.Controls.Add(
                headerPanel
            )



            '========================================================
            ' TITLE LABEL
            '========================================================

            titleLabel =
                New Label()


            titleLabel.Text =
                "3D Model Viewer"


            titleLabel.AutoSize =
                True


            titleLabel.Location =
                New DrawingPoint(
                    15,
                    9
                )


            titleLabel.Font =
                New DrawingFont(
                    "Segoe UI",
                    11.0F,
                    DrawingFontStyle.Bold
                )


            headerPanel.Controls.Add(
                titleLabel
            )



            '========================================================
            ' MODEL FILE NAME
            '========================================================

            modelNameLabel =
                New Label()


            modelNameLabel.Text =
                IOPath.GetFileName(
                    _modelPath
                )


            modelNameLabel.AutoSize =
                True


            modelNameLabel.Location =
                New DrawingPoint(
                    16,
                    35
                )


            modelNameLabel.ForeColor =
                DrawingColor.DimGray


            headerPanel.Controls.Add(
                modelNameLabel
            )



            '========================================================
            ' BOTTOM PANEL
            '========================================================

            bottomPanel =
                New Panel()


            bottomPanel.Dock =
                DockStyle.Bottom


            bottomPanel.Height =
                82


            bottomPanel.BackColor =
                DrawingColor.White


            Me.Controls.Add(
                bottomPanel
            )



            '========================================================
            ' HELP LABEL
            '========================================================

            helpLabel =
                New Label()


            helpLabel.Text =
                "Left drag: Rotate     Middle drag: Pan     Mouse wheel: Zoom"


            helpLabel.AutoSize =
                True


            helpLabel.Location =
                New DrawingPoint(
                    15,
                    8
                )


            helpLabel.ForeColor =
                DrawingColor.DimGray


            bottomPanel.Controls.Add(
                helpLabel
            )



            '========================================================
            ' FRONT BUTTON
            '========================================================

            frontButton =
                CreateButton(
                    "Front",
                    15
                )


            AddHandler frontButton.Click,
                AddressOf FrontButton_Click


            bottomPanel.Controls.Add(
                frontButton
            )



            '========================================================
            ' TOP BUTTON
            '========================================================

            topButton =
                CreateButton(
                    "Top",
                    105
                )


            AddHandler topButton.Click,
                AddressOf TopButton_Click


            bottomPanel.Controls.Add(
                topButton
            )



            '========================================================
            ' ISO BUTTON
            '========================================================

            isoButton =
                CreateButton(
                    "Iso",
                    195
                )


            AddHandler isoButton.Click,
                AddressOf IsoButton_Click


            bottomPanel.Controls.Add(
                isoButton
            )



            '========================================================
            ' FIT BUTTON
            '========================================================

            fitButton =
                CreateButton(
                    "Fit",
                    285
                )


            AddHandler fitButton.Click,
                AddressOf FitButton_Click


            bottomPanel.Controls.Add(
                fitButton
            )



            '========================================================
            ' CLOSE BUTTON
            '========================================================

            closeButton =
                New Button()


            closeButton.Text =
                "Close"


            closeButton.Size =
                New DrawingSize(
                    90,
                    30
                )


            closeButton.FlatStyle =
                FlatStyle.System


            closeButton.Anchor =
                AnchorStyles.Right Or
                AnchorStyles.Bottom


            bottomPanel.Controls.Add(
                closeButton
            )


            AddHandler closeButton.Click,
                AddressOf CloseButton_Click


            AddHandler bottomPanel.Resize,
                AddressOf BottomPanel_Resize



            '========================================================
            ' MODEL DISPLAY PANEL
            '========================================================

            modelPanel =
                New ViewerSurface()


            modelPanel.Dock =
                DockStyle.Fill


            modelPanel.BorderStyle =
                BorderStyle.FixedSingle


            modelPanel.BackColor =
                DrawingColor.White


            Me.Controls.Add(
                modelPanel
            )


            modelPanel.BringToFront()



            '========================================================
            ' MODEL PANEL EVENTS
            '========================================================

            AddHandler modelPanel.MouseDown,
                AddressOf ModelPanel_MouseDown


            AddHandler modelPanel.MouseMove,
                AddressOf ModelPanel_MouseMove


            AddHandler modelPanel.MouseUp,
                AddressOf ModelPanel_MouseUp


            AddHandler modelPanel.MouseWheel,
                AddressOf ModelPanel_MouseWheel


            AddHandler modelPanel.MouseEnter,
                AddressOf ModelPanel_MouseEnter


            AddHandler modelPanel.Resize,
                AddressOf ModelPanel_Resize


            AddHandler modelPanel.Paint,
                AddressOf ModelPanel_Paint



            '========================================================
            ' FORM EVENTS
            '========================================================

            AddHandler Me.Shown,
                AddressOf ModelViewerForm_Shown


            AddHandler Me.FormClosing,
                AddressOf ModelViewerForm_FormClosing


            BottomPanel_Resize(
                Nothing,
                EventArgs.Empty
            )

        End Sub



        '============================================================
        ' CREATE COMMON BUTTON
        '============================================================

        Private Function CreateButton(
            buttonText As String,
            leftPosition As Integer
        ) As Button


            Dim button As New Button()


            button.Text =
                buttonText


            button.Size =
                New DrawingSize(
                    80,
                    30
                )


            button.Location =
                New DrawingPoint(
                    leftPosition,
                    42
                )


            button.FlatStyle =
                FlatStyle.System


            button.BackColor =
                DrawingColor.White


            Return button

        End Function



        '============================================================
        ' FORM SHOWN
        '============================================================

        Private Sub ModelViewerForm_Shown(
            sender As Object,
            e As EventArgs
        )

            Try

                LoadModel()

            Catch ex As Exception

                MsgBox(
                    "Unable to initialize 3D viewer:" &
                    vbCrLf &
                    vbCrLf &
                    ex.ToString()
                )

            End Try

        End Sub



        '============================================================
        ' LOAD MODEL USING APPRENTICE SERVER
        '============================================================

        Private Sub LoadModel()

            Try

                '====================================================
                ' CREATE APPRENTICE SERVER
                '====================================================

                apprentice =
                    New Inventor.ApprenticeServerComponent()


                '====================================================
                ' OPTIMIZE APPRENTICE FOR 3D VIEWING
                '====================================================

                apprentice.DisplayAffinity =
                    True


                '====================================================
                ' OPEN INVENTOR MODEL
                '====================================================

                apprenticeDocument =
                    apprentice.Open(
                        _modelPath
                    )


                If apprenticeDocument Is Nothing Then

                    MsgBox(
                        "Unable to open 3D model:" &
                        vbCrLf &
                        vbCrLf &
                        _modelPath
                    )

                    Exit Sub

                End If


                '====================================================
                ' FORCE PANEL HANDLE CREATION
                '====================================================

                Dim viewerHandle As IntPtr =
                    modelPanel.Handle


                '====================================================
                ' CREATE CLIENT VIEW
                '
                ' CLIENT VIEW WILL DRAW INSIDE modelPanel
                '====================================================

                clientView =
                    apprenticeDocument.
                    ClientViews.
                    Add(
                        CLng(
                            viewerHandle.ToInt64()
                        )
                    )


                If clientView Is Nothing Then

                    MsgBox(
                        "Unable to create Inventor ClientView."
                    )

                    Exit Sub

                End If


                clientView.DisplayMode =
    Inventor.DisplayModeEnum.
    kShadedWithEdgesRendering


                '====================================================
                ' FIRST INITIALIZE / DRAW CLIENT VIEW
                '====================================================

                clientView.Update(
    False
)

                Application.DoEvents()


                '====================================================
                ' INITIAL FIT
                '====================================================

                Dim initialCamera As Inventor.Camera =
    clientView.Camera


                initialCamera.Fit()

                initialCamera.Apply()


                clientView.Update(
    False
)


                Application.DoEvents()


                '====================================================
                ' NOW APPLY ISO AFTER CLIENTVIEW IS INITIALIZED
                '====================================================

                Me.BeginInvoke(
    New MethodInvoker(
        Sub()

            SetOrientation(
                Inventor.ViewOrientationTypeEnum.
                kIsoTopRightViewOrientation
            )

        End Sub
    )
)


                modelPanel.Focus()


            Catch ex As Exception

                MsgBox(
                    "Error while loading 3D model:" &
                    vbCrLf &
                    vbCrLf &
                    _modelPath &
                    vbCrLf &
                    vbCrLf &
                    ex.ToString()
                )

            End Try

        End Sub



        '============================================================
        ' FRONT BUTTON
        '============================================================

        Private Sub FrontButton_Click(
            sender As Object,
            e As EventArgs
        )

            SetOrientation(
                Inventor.ViewOrientationTypeEnum.
                kFrontViewOrientation
            )

        End Sub



        '============================================================
        ' TOP BUTTON
        '============================================================

        Private Sub TopButton_Click(
            sender As Object,
            e As EventArgs
        )

            SetOrientation(
                Inventor.ViewOrientationTypeEnum.
                kTopViewOrientation
            )

        End Sub



        '============================================================
        ' ISO BUTTON
        '============================================================

        Private Sub IsoButton_Click(
            sender As Object,
            e As EventArgs
        )

            SetOrientation(
                Inventor.ViewOrientationTypeEnum.
                kIsoTopRightViewOrientation
            )

        End Sub



        '============================================================
        ' SET CAMERA ORIENTATION
        '============================================================

        Private Sub SetOrientation(
    orientation As Inventor.ViewOrientationTypeEnum
)

            Try

                If clientView Is Nothing Then
                    Exit Sub
                End If


                '============================================================
                ' MAKE SURE VIEW HAS BEEN INITIALIZED
                '============================================================

                clientView.Update(
            False
        )

                Application.DoEvents()


                '============================================================
                ' GET CAMERA
                '============================================================

                Dim camera As Inventor.Camera =
            clientView.Camera


                If camera Is Nothing Then

                    MsgBox(
                "Unable to get the ClientView camera."
            )

                    Exit Sub

                End If


                '============================================================
                ' SET ORIENTATION
                '============================================================

                camera.ViewOrientationType =
            orientation


                '============================================================
                ' ORTHOGRAPHIC
                '============================================================

                camera.Perspective =
            False


                '============================================================
                ' FIT COMPLETE MODEL
                '============================================================

                camera.Fit()


                '============================================================
                ' APPLY
                '============================================================

                camera.Apply()


                '============================================================
                ' REDRAW
                '============================================================

                clientView.Update(
            False
        )


            Catch ex As Exception

                MsgBox(
            "Unable to change model orientation:" &
            vbCrLf &
            vbCrLf &
            ex.ToString()
        )

            End Try

        End Sub



        '============================================================
        ' FIT BUTTON
        '============================================================

        Private Sub FitButton_Click(
            sender As Object,
            e As EventArgs
        )

            FitModel()

        End Sub



        '============================================================
        ' FIT COMPLETE MODEL
        '============================================================

        Private Sub FitModel()

            Try

                If clientView Is Nothing Then
                    Exit Sub
                End If


                Dim camera As Inventor.Camera =
                    clientView.Camera


                camera.Fit()


                camera.Apply()


                clientView.Update(
                    False
                )


            Catch ex As Exception

                MsgBox(
                    "Unable to fit model:" &
                    vbCrLf &
                    vbCrLf &
                    ex.ToString()
                )

            End Try

        End Sub



        '============================================================
        ' MOUSE ENTER
        '============================================================

        Private Sub ModelPanel_MouseEnter(
            sender As Object,
            e As EventArgs
        )

            modelPanel.Focus()

        End Sub



        '============================================================
        ' MOUSE DOWN
        '
        ' LEFT BUTTON   = ROTATE
        ' MIDDLE BUTTON = PAN
        '============================================================

        Private Sub ModelPanel_MouseDown(
            sender As Object,
            e As MouseEventArgs
        )

            If clientView Is Nothing Then
                Exit Sub
            End If


            modelPanel.Focus()


            If e.Button =
               MouseButtons.Left Then


                mouseDragging =
                    True


                currentMouseOperation =
                    "ROTATE"


                previousMousePoint =
                    e.Location


            ElseIf e.Button =
                   MouseButtons.Middle Then


                mouseDragging =
                    True


                currentMouseOperation =
                    "PAN"


                previousMousePoint =
                    e.Location

            End If

        End Sub



        '============================================================
        ' MOUSE MOVE
        '
        ' ROTATE / PAN
        '============================================================

        Private Sub ModelPanel_MouseMove(
            sender As Object,
            e As MouseEventArgs
        )

            Try

                If clientView Is Nothing Then
                    Exit Sub
                End If


                If apprentice Is Nothing Then
                    Exit Sub
                End If


                If Not mouseDragging Then
                    Exit Sub
                End If


                If currentMouseOperation = "" Then
                    Exit Sub
                End If


                Dim camera As Inventor.Camera =
                    clientView.Camera


                Dim transientGeometry As Inventor.TransientGeometry =
                    apprentice.TransientGeometry


                Dim fromPoint As Inventor.Point2d =
                    transientGeometry.CreatePoint2d(
                        CDbl(previousMousePoint.X),
                        CDbl(previousMousePoint.Y)
                    )


                Dim toPoint As Inventor.Point2d =
                    transientGeometry.CreatePoint2d(
                        CDbl(e.X),
                        CDbl(e.Y)
                    )


                '====================================================
                ' ROTATE
                '====================================================

                If currentMouseOperation =
                   "ROTATE" Then


                    camera.ComputeWithMouseInput(
                        fromPoint,
                        toPoint,
                        0,
                        Inventor.ViewOperationTypeEnum.
                            kRotateViewOperation
                    )


                    '====================================================
                    ' PAN
                    '====================================================

                ElseIf currentMouseOperation =
                       "PAN" Then


                    camera.ComputeWithMouseInput(
                        fromPoint,
                        toPoint,
                        0,
                        Inventor.ViewOperationTypeEnum.
                            kPanViewOperation
                    )

                End If


                camera.Apply()


                '====================================================
                ' TRUE = INTERACTIVE / FASTER REDRAW
                '====================================================

                clientView.Update(
                    True
                )


                previousMousePoint =
                    e.Location


            Catch ex As Exception

                mouseDragging =
                    False


                currentMouseOperation =
                    ""


                MsgBox(
                    "Error while moving 3D model:" &
                    vbCrLf &
                    vbCrLf &
                    ex.ToString()
                )

            End Try

        End Sub



        '============================================================
        ' MOUSE UP
        '============================================================

        Private Sub ModelPanel_MouseUp(
            sender As Object,
            e As MouseEventArgs
        )

            Try

                mouseDragging =
                    False


                currentMouseOperation =
                    ""


                If clientView IsNot Nothing Then

                    '================================================
                    ' FALSE = COMPLETE FINAL RENDER
                    '================================================

                    clientView.Update(
                        False
                    )

                End If


            Catch

            End Try

        End Sub



        '============================================================
        ' MOUSE WHEEL
        '
        ' ZOOM
        '============================================================

        Private Sub ModelPanel_MouseWheel(
            sender As Object,
            e As MouseEventArgs
        )

            Try

                If clientView Is Nothing Then
                    Exit Sub
                End If


                If apprentice Is Nothing Then
                    Exit Sub
                End If


                Dim camera As Inventor.Camera =
                    clientView.Camera


                Dim transientGeometry As Inventor.TransientGeometry =
                    apprentice.TransientGeometry


                Dim mousePoint As Inventor.Point2d =
                    transientGeometry.CreatePoint2d(
                        CDbl(e.X),
                        CDbl(e.Y)
                    )


                '====================================================
                ' WINDOWS MOUSE WHEEL
                '
                ' 120 DELTA = ONE CLICK
                '====================================================

                Dim wheelClicks As Integer =
                    CInt(
                        e.Delta /
                        SystemInformation.MouseWheelScrollDelta
                    )


                If wheelClicks = 0 Then
                    Exit Sub
                End If


                camera.ComputeWithMouseInput(
                    mousePoint,
                    mousePoint,
                    wheelClicks,
                    Inventor.ViewOperationTypeEnum.
                        kZoomViewOperation
                )


                camera.Apply()


                clientView.Update(
                    False
                )


            Catch ex As Exception

                MsgBox(
                    "Error while zooming 3D model:" &
                    vbCrLf &
                    vbCrLf &
                    ex.ToString()
                )

            End Try

        End Sub



        '============================================================
        ' MODEL PANEL PAINT
        '
        ' REDRAW CLIENT VIEW
        '============================================================

        Private Sub ModelPanel_Paint(
            sender As Object,
            e As PaintEventArgs
        )

            Try

                If clientView IsNot Nothing Then

                    clientView.Update(
                        False
                    )

                End If

            Catch

            End Try

        End Sub



        '============================================================
        ' MODEL PANEL RESIZE
        '============================================================

        Private Sub ModelPanel_Resize(
            sender As Object,
            e As EventArgs
        )

            Try

                If clientView IsNot Nothing Then

                    clientView.Update(
                        False
                    )

                End If

            Catch

            End Try

        End Sub



        '============================================================
        ' BOTTOM PANEL RESIZE
        '
        ' KEEP CLOSE BUTTON AT RIGHT SIDE
        '============================================================

        Private Sub BottomPanel_Resize(
            sender As Object,
            e As EventArgs
        )

            If closeButton Is Nothing Then
                Exit Sub
            End If


            closeButton.Location =
                New DrawingPoint(
                    bottomPanel.ClientSize.Width -
                    closeButton.Width -
                    15,
                    42
                )

        End Sub



        '============================================================
        ' CLOSE BUTTON
        '============================================================

        Private Sub CloseButton_Click(
            sender As Object,
            e As EventArgs
        )

            Me.Close()

        End Sub



        '============================================================
        ' CLEANUP WHEN VIEWER CLOSES
        '============================================================

        Private Sub ModelViewerForm_FormClosing(
            sender As Object,
            e As FormClosingEventArgs
        )

            mouseDragging =
                False


            currentMouseOperation =
                ""


            '========================================================
            ' CLOSE CLIENT VIEW
            '========================================================

            Try

                If clientView IsNot Nothing Then

                    clientView.Close()


                    clientView =
                        Nothing

                End If

            Catch

            End Try


            '========================================================
            ' CLOSE APPRENTICE
            '========================================================

            Try

                If apprentice IsNot Nothing Then

                    apprentice.Close()

                End If

            Catch

            End Try


            apprenticeDocument =
                Nothing


            apprentice =
                Nothing

        End Sub


    End Class


End Class