Public Class Form1
    Dim Box(9, 9) As Button
    Dim BoxFlag(9, 9) As Boolean '那些格子上有插旗子
    Dim ck As Boolean '開始了嗎
    Dim Time, Bomb As Integer '秒數 旗子數量
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        For i As Integer = 1 To 9
            For x As Integer = 1 To 9
                Box(i, x) = Me.Controls("Button" & (i - 1) * 9 + x)
                AddHandler Box(i, x).Click, AddressOf UserCk
                AddHandler Box(i, x).MouseUp, AddressOf Flag
            Next
        Next
        PictureBox1.Image = New Bitmap("22.jpg")
        PictureBox2.Image = New Bitmap("33.jpg")
        Re() : TextBox1.ReadOnly = True : TextBox2.ReadOnly = True
    End Sub

    Sub SetBomb(ByVal Now1 As Integer, ByVal Now2 As Integer) '製造炸彈 Now=現在點選的位子
        Randomize()
        Dim Key() As String = {""}
        Dim N1, N2 As Integer
        N1 = Int(Rnd() * 9 + 1) '隨機產生的y
        N2 = Int(Rnd() * 9 + 1) '隨機產生的x
        For i As Integer = 1 To Bomb  '隨機產生炸彈
            Do Until Box(N1, N2).Tag = 0 And (Now1 <> N1 Or Now2 <> N2) '不=現在點的位子 而且要要是空的 炸彈不能重複在同一個地方
                N1 = Int(Rnd() * 9 + 1)
                N2 = Int(Rnd() * 9 + 1)
            Loop
            Box(N1, N2).Tag = "9" '炸彈
        Next
        '開始編碼
        Dim SetNX() As Integer = {-1, -1, -1, 0, 0, 1, 1, 1} 'BOX(X,Y)
        Dim SetNY() As Integer = {-1, 0, 1, -1, 1, -1, 0, 1} '產生炸彈數字的陣列
        For i As Integer = 1 To 9
            For x As Integer = 1 To 9
                If Box(i, x).Tag = "9" Then '=炸彈
                    For K As Integer = 0 To 7 '增加數字
                        BombNb(i + SetNX(K), x + SetNY(K))
                    Next
                End If
            Next
        Next
        Track(Now1, Now2) '遞迴找路
    End Sub

    Sub Track(ByVal NowX As Integer, ByVal NowY As Integer) '放狀腺
        If (NowX <= 0 Or NowX >= 10) Or (NowY <= 0 Or NowY >= 10) Then Exit Sub
        If Box(NowX, NowY).Enabled = False Then Exit Sub
        If Box(NowX, NowY).Tag = 0 Then
            Box(NowX, NowY).Enabled = False : Box(NowX, NowY).Text = "" 'Box(NowX, NowY).Tag ' : Box(NowX, NowY).BackColor = Color.White
            Track(NowX - 1, NowY) '上
            Track(NowX, NowY - 1) '左
            Track(NowX, NowY + 1) '右
            Track(NowX + 1, NowY) '下
            Track(NowX - 1, NowY - 1) '左上
            Track(NowX - 1, NowY + 1) '右上
            Track(NowX + 1, NowY - 1) '左下
            Track(NowX + 1, NowY + 1) '右下
        ElseIf Box(NowX, NowY).Tag <> "9" Then
            Box(NowX, NowY).Enabled = False ': Box(NowX, NowY).BackColor = Color.White
            Box(NowX, NowY).Text = Box(NowX, NowY).Tag
        End If
    End Sub

    Sub BombNb(ByVal X As Integer, ByVal Y As Integer) '算炸彈旁的數字 (累計)
        If (X > 0 And X < 10) And (Y > 0 And Y < 10) Then
            If Box(X, Y).Tag <> "9" Then Box(X, Y).Tag += 1
        End If
    End Sub

    Sub UserCk(ByVal sender As System.Object, ByVal e As System.EventArgs) '踩地雷
        Dim Now1, Now2 As Integer
        Now1 = (Val(Mid(CType(sender, Button).Name, 7)) - 1) \ 9 + 1 '現在點選的位子
        Now2 = Val(Mid(CType(sender, Button).Name, 7) - 1) Mod 9 + 1 '現在點選的位子
        If BoxFlag(Now1, Now2) = False Then
            If ck = False Then '第一次點選
                ck = 1
                SetBomb(Now1, Now2)
            ElseIf ck = True Then
                If Box(Now1, Now2).Tag = 0 Then '踩到空的地方
                    Track(Now1, Now2) '放狀腺
                ElseIf Box(Now1, Now2).Tag = "9" Then '踩到炸彈
                    Timer1.Enabled = False
                    For i As Integer = 1 To 9 '全部關閉
                        For x As Integer = 1 To 9
                            Box(i, x).Image = Nothing
                        Next
                        For x As Integer = 1 To 9
                            Box(i, x).Enabled = False
                            If Box(i, x).Tag = "9" Then Box(i, x).Image = New Bitmap("9.png") '顯示炸彈圖
                        Next
                    Next
                    If MsgBox("GameOver" & vbNewLine & "是否要重新開始？", 1) = vbOK Then
                        Re()
                    End If
                    Exit Sub
                Else : Box(Now1, Now2).Text = Box(Now1, Now2).Tag : Box(Now1, Now2).Enabled = False
                End If
            End If
        End If
        Read()
    End Sub

    Sub Re() '重製
        ck = False : Timer1.Enabled = True
        For i As Integer = 1 To 9
            For x As Integer = 1 To 9
                Box(i, x).Tag = 0
                Box(i, x).Text = ""
                Box(i, x).Enabled = True
                Box(i, x).Image = Nothing
                Box(i, x).BackColor = Color.SkyBlue
                BoxFlag(i, x) = False
            Next
        Next
        Bomb = 10
        TextBox1.Text = 0 : TextBox2.Text = Bomb : Time = 0
    End Sub

    Sub Read() '塗顏色 偵測結束了沒
        Dim fr As Integer '場上還有幾個格子
        For i As Integer = 1 To 9
            For x As Integer = 1 To 9
                If Box(i, x).Enabled = False Then Box(i, x).BackColor = Color.White Else fr += 1
                If Box(i, x).Text = "1" Then
                    Box(i, x).BackColor = Color.FromArgb(236, 224, 200)
                ElseIf Box(i, x).Text = "2" Then
                    Box(i, x).BackColor = Color.FromArgb(238, 228, 128)
                ElseIf Box(i, x).Text = "3" Then
                    Box(i, x).BackColor = Color.FromArgb(242, 177, 112)
                ElseIf Box(i, x).Text = "4" Then
                    Box(i, x).BackColor = Color.FromArgb(236, 141, 83)
                ElseIf Box(i, x).Text = "5" Then
                    Box(i, x).BackColor = Color.FromArgb(245, 124, 95)
                ElseIf Box(i, x).Text = "6" Then
                    Box(i, x).BackColor = Color.FromArgb(233, 89, 55)
                ElseIf Box(i, x).Text = "7" Then
                    Box(i, x).BackColor = Color.FromArgb(243, 127, 107)
                ElseIf Box(i, x).Text = "8" Then
                    Box(i, x).BackColor = Color.FromArgb(241, 208, 75)
                End If
            Next
        Next
        If fr = 10 Then '場上格子跟炸彈一樣多
            Timer1.Enabled = False
            For i As Integer = 1 To 9 '全部關閉
                For x As Integer = 1 To 9
                    Box(i, x).Enabled = False
                    If Box(i, x).Tag = "9" Then Box(i, x).Image = New Bitmap("9.png") '顯示炸彈圖
                Next
            Next
            If MsgBox("恭喜破關，一共花了 " & Time & " 秒，是否要重新開始？", 1) = vbOK Then Re()
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick '時間
        If ck = True Then 'ck=true =開始
            Time += 1
        End If
        TextBox1.Text = Time
    End Sub

    Private Sub 重新開始ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 重新開始ToolStripMenuItem.Click
        Re()
    End Sub

    Sub Flag(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) '插旗子
        Dim Now1, Now2 As Integer
        Now1 = (Val(Mid(CType(sender, Button).Name, 7)) - 1) \ 9 + 1 '現在點選的位子
        Now2 = Val(Mid(CType(sender, Button).Name, 7) - 1) Mod 9 + 1 '現在點選的位子
        If e.Button = Windows.Forms.MouseButtons.Right Then '插旗子 如果點的是滑鼠右鍵
            If BoxFlag(Now1, Now2) = False Then
                Box(Now1, Now2).Image = New Bitmap("1.png") : BoxFlag(Now1, Now2) = True
                Bomb -= 1 '旗子數量
            Else : Box(Now1, Now2).Image = Nothing : BoxFlag(Now1, Now2) = False : Bomb += 1
            End If
        End If
        TextBox2.Text = Bomb
    End Sub

    Private Sub 結束ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        End
    End Sub

    Private Sub 結束ToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 結束ToolStripMenuItem.Click
        End
    End Sub

    Private Sub 設定ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 設定ToolStripMenuItem.Click
        Try
            Dim n As Integer = InputBox("請輸入炸彈數量")
            If n < 10 Or n > 80 Then Throw New Exception Else Re() : Bomb = n : TextBox2.Text = Bomb
        Catch ex As Exception
            MsgBox("輸入錯誤")
        End Try
    End Sub
End Class
