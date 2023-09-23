#If VBA7 And Win64 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#Else
    Private Declare  Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
    Private Declare  Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#End If

Dim axis As String
Dim value As Single

Const origin_x = 350
Const origin_y = 400
Const margin = 250
Const originF1_x = origin_x + margin + 50
Const originF1_y = origin_y + margin

Const originF2_x = origin_x - margin + 100
Const originF2_y = origin_y + margin

Const originF3_x = origin_x + 50
Const originF3_y = origin_y - margin
Dim margin_s As Single
Const origin_arrow_width = 50
Const origin_arrow_height = 120

Dim shape As shape
Dim flag As Single

Private Sub common_start()
    ' 画面描画をオフ 見栄えと高速化
    Application.ScreenUpdating = False
End Sub

Private Sub common_end()
    ' 画面描画オン
    Application.ScreenUpdating = True
End Sub


Private Sub ResetShape(ByVal shp As shape, Optional ByVal x As Single, Optional ByVal y As Single, _
Optional ByVal RotationX As Single, Optional ByVal RotationY As Single, Optional ByVal RotationZ As Single)
    With shp.ThreeD
        .RotationX = RotationX
        .RotationY = RotationY
        .RotationZ = RotationZ
    End With
    
    If Not IsMissing(x) Then shp.Left = x
    If Not IsMissing(y) Then shp.Top = y
End Sub


' 角度の初期化
Sub initialize()
    common_start
    
    ' 各シェイプに対してリセット処理を適用
    For Each shp In ActiveSheet.Shapes
    Debug.Print shp.Name
        ' シェイプの名前に応じてリセット
        Select Case Left(shp.Name, 5)
            Case "core"
                ResetShape shp, origin_x, origin_y, 0, 0, 0
            Case "face1"
                ResetShape shp, , , 0, 0, 0
            Case "face2"
                ResetShape shp, , , 270, 0, 0
            Case "face3"
                ResetShape shp, , , 0, 270, 0
        End Select
    Next
    
    ' 回転情報の取得
    Call get_rotation
    
    ' ロックの解除
    release
    
    common_end
End Sub


Sub rotate_common(axis, minus As Integer, bitFlag As Integer)
    common_start
    
    ' 大小
    If bitFlag = 1 Then
        value = 3
    Else
        value = 15
    End If
    
    '正負
    If minus = -1 Then
        value = value * -1
    End If
    
    ' Shift同時で精緻モード
    If GetKeyState(vbKeyShift) < 0 Then
        value = value / 3
    End If

    For Each shp In ActiveSheet.Shapes
        If Left(shp.Name, 4) = "face" Or shp.Name = "core" Then

            Select Case axis
                Case "h"
                    shp.ThreeD.IncrementRotationHorizontal value
                Case "v"
                    shp.ThreeD.IncrementRotationVertical value
                Case "x"
                    shp.ThreeD.IncrementRotationX value
                Case "y"
                    shp.ThreeD.IncrementRotationY value
                Case "z"
                    shp.ThreeD.IncrementRotationZ value
            End Select
        End If
    Next
    
    get_rotation
       
    rotate_adjust_to_lock
    
    ' 位置補正具合を見るにはここを有効に
    'ActiveSheet.Shapes.Range(Array("face1r", "face1l", "face1t", "face1b", "face2r", "face2l", "face2b", "face2t", "face3r", "face3l", "face3b", "face3t")).Select
    
    common_end
End Sub

Sub rotate_x()
    Call rotate_common("x", 1, 0)
End Sub

Sub rotate_x_minus()
    Call rotate_common("x", -1, 0)
End Sub

Sub rotate_y()
    Call rotate_common("y", 1, 0)
End Sub

Sub rotate_y_minus()
    Call rotate_common("y", -1, 0)
End Sub

Sub rotate_vertical()
    Call rotate_common("v", 1, 0)
End Sub
Sub rotate_vertical_bit()
    Call rotate_common("v", 1, 1)
End Sub

Sub rotate_vertical_minus()
    Call rotate_common("v", -1, 0)
End Sub

Sub rotate_vertical_minus_bit()
    Call rotate_common("v", -1, 1)
End Sub

Sub rotate_horizontal()
    Call rotate_common("h", 1, 0)
End Sub

Sub rotate_horizontal_bit()
    Call rotate_common("h", 1, 1)
End Sub

Sub rotate_horizontal_minus()
    Call rotate_common("h", -1, 0)
End Sub

Sub rotate_horizontal_minus_bit()
    Call rotate_common("h", -1, 1)
End Sub
Sub rotate_z()
    Call rotate_common("z", 1, 0)
End Sub

Sub rotate_z_bit()
    Call rotate_common("z", 1, 1)
End Sub

Sub rotate_z_minus()
    Call rotate_common("z", -1, 0)
End Sub

Sub rotate_z_minus_bit()
    Call rotate_common("z", -1, 1)
End Sub


Sub rotate_common_sub(axis, value)
    
    For Each shp In ActiveSheet.Shapes
        If Left(shp.Name, 4) = "face" Or shp.Name = "core" Then
            Select Case axis
                Case "h"
                    shp.ThreeD.IncrementRotationHorizontal value
                Case "v"
                    shp.ThreeD.IncrementRotationVertical value
                Case "x"
                    shp.ThreeD.IncrementRotationX value
                Case "y"
                    shp.ThreeD.IncrementRotationY value
                Case "z"
                    shp.ThreeD.IncrementRotationZ value
            End Select
        End If
    Next
    
    Call get_rotation
    
End Sub



Sub set_lock1()
     If Not Range("degree1") = "-" Then
        common_start
        
        Range("flag_lock") = 1
        Range("lock_degree1") = Range("degree1")
        Range("lock_degree2").ClearContents
        Range("lock_degree3").ClearContents
        
        lock_btn_display
        
        common_end
    End If
End Sub

Sub set_lock2()
    If Not Range("degree2") = "-" Then
        common_start
    
        Range("flag_lock") = 2
        Range("lock_degree1").ClearContents
        Range("lock_degree2") = Range("degree2")
        Range("lock_degree3").ClearContents
        
        lock_btn_display
        
        common_end
    End If
End Sub

Sub set_lock3()
    If Not Range("degree3") = "-" Then
        common_start
    
        Range("flag_lock") = 3
        Range("lock_degree1").ClearContents
        Range("lock_degree2").ClearContents
        Range("lock_degree3") = Range("degree3")
        
        lock_btn_display
    
        common_end
    End If
End Sub

Sub rotate_adjust_to_lock()
    ' ロックフラグの値を取得
    flagLock = Range("flag_lock")
    
    ' ロックフラグの値に応じて、回転調整を行う
    If flagLock >= 1 And flagLock <= 3 Then
        degreeIndex = "degree" & flagLock
        lockDegreeIndex = "lock_degree" & flagLock
        Call rotate_common_sub("z", Round(Range(degreeIndex) - Range(lockDegreeIndex), 1))
    End If
End Sub


Sub lock_btn_display()
    Dim selectedButton As String
    Dim flagLock As Integer
    
    ' flag_lockの値を取得
    flagLock = Range("flag_lock").value
    
    ' flag_lockの値が1から3の間の場合のみ処理を実行
    If flagLock >= 1 And flagLock <= 3 Then
        ' 他のボタンの色をリセットし、適切なボタンを選択
        For i = 1 To 3
            Dim buttonName As String
            buttonName = "btn_lock" & i
            
            If i = flagLock Then
                ' 選択されたボタンの名前を保存
                selectedButton = buttonName
            Else
                ' 他のボタンの色をリセット
                Call reset_gray(buttonName)
            End If
        Next i
        
        ' 選択されたボタンの色をグレーに設定
        With ActiveSheet.Shapes.Range(Array(selectedButton))
            .Fill.ForeColor.RGB = RGB(221, 221, 221)
            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(95, 95, 95)
            .Line.ForeColor.RGB = RGB(192, 192, 192)
        End With
        
        ' lock_releaseシェイプを最前面に移動し、可視化
        With ActiveSheet.Shapes.Range(Array("lock_release"))
            .Visible = msoTrue
            .ZOrder msoBringToFront
        End With
    End If
End Sub


Sub release()
    common_start
    
    ' ロック関連のセルの内容をクリア
    Range("flag_lock").ClearContents
    Range("lock_degree1").ClearContents
    Range("lock_degree2").ClearContents
    Range("lock_degree3").ClearContents

    ' 目隠しシェイプを非表示にする
    If ActiveSheet.Shapes.Range(Array("lock_release")).Visible Then
        ActiveSheet.Shapes.Range(Array("lock_release")).Visible = msoFalse
    End If
    
    ' 各ロックボタンの色をリセット
    reset_gray ("btn_lock1")
    reset_gray ("btn_lock2")
    reset_gray ("btn_lock3")

    common_end
End Sub

Sub reset_gray(object_name)
    Set shape = ActiveSheet.Shapes(object_name)
    
    With shape
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Line.ForeColor.ObjectThemeColor = msoThemeColorText1
        .Line.ForeColor.Brightness = 0.5
    End With
End Sub


Sub arrow_depth_up()
    common_start

    For Each shp In ActiveSheet.Shapes
        If Left(shp.Name, 4) = "face" Then
            shp.ThreeD.Depth = shp.ThreeD.Depth + 5
        End If
    Next
    
    get_rotation

    common_end
End Sub

Sub arrow_depth_down()
    common_start

    For Each shp In ActiveSheet.Shapes
        If Left(shp.Name, 4) = "face" Then
            If shp.ThreeD.Depth < 5 Then
                shp.ThreeD.Depth = 0
            Else
                shp.ThreeD.Depth = shp.ThreeD.Depth - 5
            End If
        End If
    Next
    
    get_rotation

    common_end
End Sub

Sub arrow_width_up()
    common_start

    For Each shp In ActiveSheet.Shapes
        If Left(shp.Name, 4) = "face" Then
            Select Case Right(shp.Name, 1)
                Case "t", "b"
                    shp.Width = shp.Width + 5
                Case "l", "r"
                    shp.Height = shp.Height + 5
            End Select
        End If
    Next
    
    get_rotation
    
    common_end
End Sub

Sub arrow_width_down()
    common_start

    For Each shp In ActiveSheet.Shapes
        If Left(shp.Name, 4) = "face" Then
            Select Case Right(shp.Name, 1)
                Case "t", "b"
                    If shp.Width > 5 Then
                        shp.Width = shp.Width - 5
                    End If
                Case "l", "r"
                    If shp.Height > 5 Then
                        shp.Height = shp.Height - 5
                    End If
            End Select
        End If
    Next
    
    get_rotation
    
    common_end
End Sub

Sub arrow_length_up()
    common_start

    For Each shp In ActiveSheet.Shapes
        If Left(shp.Name, 4) = "face" Then
            Select Case Right(shp.Name, 1)
                Case "t", "b"
                    shp.Height = shp.Height + 5
                Case "l", "r"
                    shp.Width = shp.Width + 5
            End Select
        End If
    Next
    
    get_rotation
        
    common_end
End Sub


Sub arrow_length_down()
    common_start

    For Each shp In ActiveSheet.Shapes
        If Left(shp.Name, 4) = "face" Then
            Select Case Right(shp.Name, 1)
                Case "t", "b"
                    If shp.Height > 5 Then
                        shp.Height = shp.Height - 5
                    End If
                Case "l", "r"
                    If shp.Width > 5 Then
                        shp.Width = shp.Width - 5
                    End If
            End Select
        End If
    Next
    
    get_rotation
    
    common_end
End Sub

Sub arrow_head_up()
    common_start

    For Each shp In ActiveSheet.Shapes
        If Left(shp.Name, 4) = "face" Then
            If shp.AutoShapeType = 33 Or shp.AutoShapeType = 35 Or shp.AutoShapeType = 36 Then
                shp.Adjustments.Item(2) = shp.Adjustments.Item(2) + 0.1
            End If
        End If
    Next
    
    get_rotation
    
    common_end
End Sub

Sub arrow_head_down()
    common_start

    For Each shp In ActiveSheet.Shapes
        If Left(shp.Name, 4) = "face" Then
            If shp.AutoShapeType = 33 Or shp.AutoShapeType = 35 Or shp.AutoShapeType = 36 Then
                shp.Adjustments.Item(2) = shp.Adjustments.Item(2) - 0.1
            End If
        End If
    Next
    
    get_rotation
    
    common_end
End Sub

Sub arrow_body_up()
    common_start

    For Each shp In ActiveSheet.Shapes
        If Left(shp.Name, 4) = "face" Then
            If shp.AutoShapeType = 33 Or shp.AutoShapeType = 35 Or shp.AutoShapeType = 36 Then
                shp.Adjustments.Item(1) = shp.Adjustments.Item(1) + 0.1
            End If
        End If
    Next
    
    get_rotation
    
    common_end
End Sub

Sub arrow_body_down()
    common_start

    For Each shp In ActiveSheet.Shapes
        If Left(shp.Name, 4) = "face" Then
            If shp.AutoShapeType = 33 Or shp.AutoShapeType = 35 Or shp.AutoShapeType = 36 Then
                shp.Adjustments.Item(1) = shp.Adjustments.Item(1) - 0.1
            End If
        End If
    Next
    
    get_rotation
    
    common_end
End Sub


Sub initialize_arrow_size()
    common_start

    For Each shp In ActiveSheet.Shapes
        If Left(shp.Name, 4) = "face" Then
            shp.ThreeD.Depth = 5
            Select Case Right(shp.Name, 1)
                Case "t", "b"
                    shp.Width = origin_arrow_width
                    shp.Height = origin_arrow_height
                Case "l", "r"
                    shp.Width = origin_arrow_height
                    shp.Height = origin_arrow_width
            End Select
            
            If shp.AutoShapeType = 33 Or shp.AutoShapeType = 35 Or shp.AutoShapeType = 36 Then
                shp.Adjustments.Item(1) = 0.6
                shp.Adjustments.Item(2) = 1
            End If
        End If
    Next
    
    get_rotation
    
    common_end
End Sub


Sub get_rotation()

    Range("rX") = ActiveSheet.Shapes.Range(Array("core")).ThreeD.RotationX
    Range("rY") = ActiveSheet.Shapes.Range(Array("core")).ThreeD.RotationY
    Range("rZ") = ActiveSheet.Shapes.Range(Array("core")).ThreeD.RotationZ
    
    Range("rX") = Round(Range("rX"), 1)
    Range("rY") = Round(Range("rY"), 1)
    Range("rZ") = Round(Range("rZ"), 1)
    
    Range("arrow_width") = ActiveSheet.Shapes.Range(Array("face1t")).Width
    Range("arrow_length") = ActiveSheet.Shapes.Range(Array("face1t")).Height
    Range("arrow_depth") = ActiveSheet.Shapes.Range(Array("face1t")).ThreeD.Depth
    
    adjust_position
End Sub



Sub adjust_position()

    margin_s = Range("arrow_width") * 0.4

    Dim arrow_names As Variant
    arrow_names = Array("face1t", "face1b", "face1r", "face1l", "face2t", "face2b", "face2r", "face2l", "face3t", "face3b", "face3r", "face3l")
    
    For Each Name In arrow_names
        
        ' 基準位置
        Select Case Left(ActiveSheet.Shapes.Range(Array(Name)).Name, 5)
            Case "face1"
                left1 = originF1_x + Range("offsetx1")
                top1 = originF1_y + Range("offsety1")
            Case "face2"
                left1 = originF2_x + Range("offsetx2")
                top1 = originF2_y + Range("offsety2")
            Case "face3"
                left1 = originF3_x + Range("offsetx3")
                top1 = originF3_y + Range("offsety3")
        End Select
        
        
        ' 矢印中心合わせ
        Select Case Right(ActiveSheet.Shapes.Range(Array(Name)).Name, 1)
            Case "t", "b"
                left2 = -Range("arrow_width") * 0.5
                top2 = -Range("arrow_length") * 0.5
            Case "l", "r"
                left2 = -Range("arrow_length") * 0.5
                top2 = -Range("arrow_width") * 0.5
        End Select
    
        ' 十字補正
        Select Case Right(ActiveSheet.Shapes.Range(Array(Name)).Name, 2)
            Case "1t"
                left3 = (Range("arrow_length") * 0.5 + margin_s) * Range("hoseix2")
                top3 = -(Range("arrow_length") * 0.5 + margin_s) * Range("hoseiy2")
            Case "1b"
                left3 = -(Range("arrow_length") * 0.5 + margin_s) * Range("hoseix2")
                top3 = (Range("arrow_length") * 0.5 + margin_s) * Range("hoseiy2")
            Case "1r"
                left3 = (Range("arrow_length") * 0.5 + margin_s) * Range("hoseix1")
                top3 = -(Range("arrow_length") * 0.5 + margin_s) * Range("hoseiy1")
            Case "1l"
                left3 = -(Range("arrow_length") * 0.5 + margin_s) * Range("hoseix1")
                top3 = (Range("arrow_length") * 0.5 + margin_s) * Range("hoseiy1")
        
            Case "2t"
                left3 = (Range("arrow_length") * 0.5 + margin_s) * Range("hoseix2")
                top3 = -(Range("arrow_length") * 0.5 + margin_s) * Range("hoseiy2")
            Case "2b"
                left3 = -(Range("arrow_length") * 0.5 + margin_s) * Range("hoseix2")
                top3 = (Range("arrow_length") * 0.5 + margin_s) * Range("hoseiy2")
            Case "2r"
                left3 = (Range("arrow_length") * 0.5 + margin_s) * Range("hoseix3")
                top3 = -(Range("arrow_length") * 0.5 + margin_s) * Range("hoseiy3")
            Case "2l"
                left3 = -(Range("arrow_length") * 0.5 + margin_s) * Range("hoseix3")
                top3 = (Range("arrow_length") * 0.5 + margin_s) * Range("hoseiy3")
        
            Case "3t"
                left3 = -(Range("arrow_length") * 0.5 + margin_s) * Range("hoseix3")
                top3 = (Range("arrow_length") * 0.5 + margin_s) * Range("hoseiy3")
            Case "3b"
                left3 = (Range("arrow_length") * 0.5 + margin_s) * Range("hoseix3")
                top3 = -(Range("arrow_length") * 0.5 + margin_s) * Range("hoseiy3")
            Case "3r"
                left3 = (Range("arrow_length") * 0.5 + margin_s) * Range("hoseix1")
                top3 = -(Range("arrow_length") * 0.5 + margin_s) * Range("hoseiy1")
            Case "3l"
                left3 = -(Range("arrow_length") * 0.5 + margin_s) * Range("hoseix1")
                top3 = (Range("arrow_length") * 0.5 + margin_s) * Range("hoseiy1")
        End Select
        
        ActiveSheet.Shapes.Range(Array(Name)).Left = left1 + left2 + left3
        ActiveSheet.Shapes.Range(Array(Name)).Top = top1 + top2 + top3
    Next

End Sub

Sub face1_right()
    common_start

    Range("offsetx1") = Range("offsetx1") + 20
    adjust_position
    
    common_end
End Sub

Sub face1_left()
    common_start

    Range("offsetx1") = Range("offsetx1") - 20
    adjust_position
    
    common_end
End Sub

Sub face1_up()
    common_start

    Range("offsety1") = Range("offsety1") - 20
    adjust_position
    
    common_end
End Sub

Sub face1_down()
    common_start

    Range("offsety1") = Range("offsety1") + 20
    adjust_position
    
    common_end
End Sub

Sub face1_reset()
    common_start
    
    Range("offsetx1") = 0
    Range("offsety1") = 0
    adjust_position
    
    common_end
End Sub


Sub face2_right()
    common_start

    Range("offsetx2") = Range("offsetx2") + 20
    adjust_position
    
    common_end
End Sub

Sub face2_left()
    common_start

    Range("offsetx2") = Range("offsetx2") - 20
    adjust_position
    
    common_end
End Sub

Sub face2_up()
    common_start

    Range("offsety2") = Range("offsety2") - 20
    adjust_position
    
    common_end
End Sub

Sub face2_down()
    common_start

    Range("offsety2") = Range("offsety2") + 20
    adjust_position
    
    common_end
End Sub

Sub face2_reset()
    common_start
    
    Range("offsetx2") = 0
    Range("offsety2") = 0
    adjust_position
    
    common_end
End Sub


Sub face3_right()
    common_start

    Range("offsetx3") = Range("offsetx3") + 20
    adjust_position
    
    common_end
End Sub

Sub face3_left()
    common_start

    Range("offsetx3") = Range("offsetx3") - 20
    adjust_position
    
    common_end
End Sub

Sub face3_up()
    common_start

    Range("offsety3") = Range("offsety3") - 20
    adjust_position
    
    common_end
End Sub

Sub face3_down()
    common_start

    Range("offsety3") = Range("offsety3") + 20
    adjust_position
    
    common_end
End Sub

Sub face3_reset()
    common_start
    
    Range("offsetx3") = 0
    Range("offsety3") = 0
    adjust_position
    
    common_end
End Sub




Sub switch_shapetype()
    common_start
    
    If ActiveSheet.Shapes.Range(Array("face1t")).AutoShapeType = 1 Then
        flag = 1
    Else
        flag = 0
    End If
    
    For Each shp In ActiveSheet.Shapes
        If Left(shp.Name, 4) = "face" Then
            If flag = 1 Then
                ' それぞれを矢印に戻す
                Select Case Right(shp.Name, 1)
                    Case "r", "l"
                        shp.AutoShapeType = 33
                    Case "t"
                        shp.AutoShapeType = 35
                    Case "b"
                        shp.AutoShapeType = 36
                End Select
                
                shp.Adjustments.Item(1) = 0.6
                shp.Adjustments.Item(2) = 1
            Else
                shp.AutoShapeType = 1
            End If
        End If
    Next
        
    common_end
End Sub

Sub switch_face()
    common_start
    
        For Each shp In ActiveSheet.Shapes
            If Left(shp.Name, 4) = "face" Then
                Select Case Left(shp.Name, 5)
                    Case "face1"
                        shp.ThreeD.RotationX = shp.ThreeD.RotationX + 180
                    Case "face2"
                        shp.ThreeD.RotationX = shp.ThreeD.RotationX + 180
                    Case "face3"
                        shp.ThreeD.RotationX = shp.ThreeD.RotationX + 180
                        shp.ThreeD.RotationZ = shp.ThreeD.RotationZ + 180
                End Select
            End If
        Next
        
    common_end
End Sub

Sub switch_transparency()
    common_start

    flag = ActiveSheet.Shapes.Range(Array("face1t")).Fill.Transparency
   
    For Each shp In ActiveSheet.Shapes
        If Left(shp.Name, 4) = "face" Then
            Select Case flag
                Case 0
                    shp.Fill.Transparency = 0.3
                Case 0.3
                    shp.Fill.Transparency = 1
                Case Else
                    shp.Fill.Transparency = 0
            End Select
        End If
    Next

    common_end
End Sub


Sub switch_ContourWidth()
    common_start

    flag = ActiveSheet.Shapes.Range(Array("face1t")).ThreeD.ContourWidth

    For Each shp In ActiveSheet.Shapes
        If Left(shp.Name, 4) = "face" Then
            If flag = 0 Then
                shp.ThreeD.ContourWidth = 1
            Else
                shp.ThreeD.ContourWidth = 0
            End If
        End If
    Next

        
    common_end
End Sub

Sub switch_color()
    common_start
      
    If ActiveSheet.Shapes.Range(Array("face1t")).Fill.ForeColor.RGB = RGB(59, 255, 82) Then
        flag = 1
    Else
        flag = 0
    End If
    
    
    For Each shp In ActiveSheet.Shapes
        If Left(shp.Name, 4) = "face" Then
            If flag = 1 Then
                shp.Fill.ForeColor.RGB = RGB(221, 221, 221)
            Else
                Select Case Right(shp.Name, 2)
                    Case "1l", "1r", "3r", "3l"
                        shp.Fill.ForeColor.RGB = RGB(255, 113, 113)
                    Case "1t", "1b", "2t", "2b"
                        shp.Fill.ForeColor.RGB = RGB(59, 255, 82)
                    Case "2l", "2r", "3t", "3b"
                        shp.Fill.ForeColor.RGB = RGB(116, 178, 252)
                End Select
            End If
        End If
    Next
    
    common_end
End Sub


Sub switch_core_visible()

    Set shape = ActiveSheet.Shapes("core")

    If shape.Fill.Transparency = 0 Then
        shape.Fill.Transparency = 1
        shape.TextFrame2.TextRange.Font.Fill.Transparency = 1
        shape.ThreeD.ContourWidth = 0
    Else
        shape.Fill.Transparency = 0
        shape.TextFrame2.TextRange.Font.Fill.Transparency = 0
        shape.ThreeD.ContourWidth = 2.5
    End If
End Sub

Sub to_back()
    If VarType(Selection) = vbObject Then
        Selection.ShapeRange.ZOrder msoSendToBack
    End If
End Sub

Sub to_front()
    If VarType(Selection) = vbObject Then
        Selection.ShapeRange.ZOrder msoBringToFront
    End If
End Sub
