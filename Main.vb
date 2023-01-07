Option Explicit On 
Option Strict On
Imports System.Text.RegularExpressions

Public Class frmMain
    Inherits System.Windows.Forms.Form

#Region " Windows フォーム デザイナで生成されたコード "

    Public Sub New()
        MyBase.New()

        ' この呼び出しは Windows フォーム デザイナで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後に初期化を追加します。

    End Sub

    ' Form は、コンポーネント一覧に後処理を実行するために dispose をオーバーライドします。
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' Windows フォーム デザイナで必要です。
    Private components As System.ComponentModel.IContainer

    ' メモ : 以下のプロシージャは、Windows フォーム デザイナで必要です。
    'Windows フォーム デザイナを使って変更してください。  
    ' コード エディタを使って変更しないでください。
    Friend WithEvents lstAddress As System.Windows.Forms.TextBox
    Friend WithEvents tmrFetch As System.Windows.Forms.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.lstAddress = New System.Windows.Forms.TextBox
        Me.tmrFetch = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'lstAddress
        '
        Me.lstAddress.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstAddress.Location = New System.Drawing.Point(0, 0)
        Me.lstAddress.Multiline = True
        Me.lstAddress.Name = "lstAddress"
        Me.lstAddress.Size = New System.Drawing.Size(292, 266)
        Me.lstAddress.TabIndex = 1
        Me.lstAddress.Text = ""
        '
        'tmrFetch
        '
        Me.tmrFetch.Enabled = True
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 12)
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.lstAddress)
        Me.Name = "frmMain"
        Me.Text = "Mail address getter"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub tmrFetch_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrFetch.Tick
        Static ssave As String
        Dim oclip As IDataObject
        Try
            'クリップボードにHTML形式が在って
            oclip = Clipboard.GetDataObject()
            If oclip.GetDataPresent(DataFormats.Html) Then

                'まだ検索していなかったら
                Dim shtml As String
                shtml = CType(oclip.GetData(DataFormats.Html), String)
                If ssave <> shtml Then
                    ssave = shtml

                    '全てのメールアドレスを検索
                    Dim fexist As Boolean = False
                    Dim oreg As Regex
                    oreg = New Regex("\b[-\w.]+@[-\w.]+\.[-\w]+\b")
                    Dim omatch As Match
                    omatch = oreg.Match(shtml)
                    lstAddress.Clear()
                    While omatch.Success
                        lstAddress.Text = lstAddress.Text & omatch.Value & vbCrLf
                        omatch = omatch.NextMatch()
                        fexist = True
                    End While
                    omatch = Nothing
                    oreg = Nothing
                    If fexist Then
                    Else
                        lstAddress.Text = "メールアドレスは見つかりません。"
                    End If
                End If
            End If
            oclip = Nothing
        Catch ex As Exception
        End Try
    End Sub

End Class
