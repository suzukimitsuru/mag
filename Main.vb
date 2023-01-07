Option Explicit On 
Option Strict On
Imports System.Text.RegularExpressions

Public Class frmMain
    Inherits System.Windows.Forms.Form

#Region " Windows �t�H�[�� �f�U�C�i�Ő������ꂽ�R�[�h "

    Public Sub New()
        MyBase.New()

        ' ���̌Ăяo���� Windows �t�H�[�� �f�U�C�i�ŕK�v�ł��B
        InitializeComponent()

        ' InitializeComponent() �Ăяo���̌�ɏ�������ǉ����܂��B

    End Sub

    ' Form �́A�R���|�[�l���g�ꗗ�Ɍ㏈�������s���邽�߂� dispose ���I�[�o�[���C�h���܂��B
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' Windows �t�H�[�� �f�U�C�i�ŕK�v�ł��B
    Private components As System.ComponentModel.IContainer

    ' ���� : �ȉ��̃v���V�[�W���́AWindows �t�H�[�� �f�U�C�i�ŕK�v�ł��B
    'Windows �t�H�[�� �f�U�C�i���g���ĕύX���Ă��������B  
    ' �R�[�h �G�f�B�^���g���ĕύX���Ȃ��ł��������B
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
            '�N���b�v�{�[�h��HTML�`�����݂���
            oclip = Clipboard.GetDataObject()
            If oclip.GetDataPresent(DataFormats.Html) Then

                '�܂��������Ă��Ȃ�������
                Dim shtml As String
                shtml = CType(oclip.GetData(DataFormats.Html), String)
                If ssave <> shtml Then
                    ssave = shtml

                    '�S�Ẵ��[���A�h���X������
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
                        lstAddress.Text = "���[���A�h���X�͌�����܂���B"
                    End If
                End If
            End If
            oclip = Nothing
        Catch ex As Exception
        End Try
    End Sub

End Class
