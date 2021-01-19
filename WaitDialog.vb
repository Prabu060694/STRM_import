Public Class WaitDialog
    Inherits System.Windows.Forms.Form
    Private bAborting As Boolean = False   ' ���~�t���O
    Private bShowing As Boolean = False   ' �_�C�A���O�\�����t���O

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
    Public WithEvents progBarMeter As System.Windows.Forms.ProgressBar
    Public WithEvents labelProgress As System.Windows.Forms.Label
    Public WithEvents labelSubMsg As System.Windows.Forms.Label
    Public WithEvents labelMainMsg As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.progBarMeter = New System.Windows.Forms.ProgressBar
        Me.labelProgress = New System.Windows.Forms.Label
        Me.labelSubMsg = New System.Windows.Forms.Label
        Me.labelMainMsg = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'progBarMeter
        '
        Me.progBarMeter.Location = New System.Drawing.Point(13, 64)
        Me.progBarMeter.Name = "progBarMeter"
        Me.progBarMeter.Size = New System.Drawing.Size(408, 23)
        Me.progBarMeter.TabIndex = 18
        '
        'labelProgress
        '
        Me.labelProgress.Font = New System.Drawing.Font("�l�r �S�V�b�N", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.labelProgress.Location = New System.Drawing.Point(13, 40)
        Me.labelProgress.Name = "labelProgress"
        Me.labelProgress.Size = New System.Drawing.Size(408, 16)
        Me.labelProgress.TabIndex = 17
        Me.labelProgress.Text = "�f�[�^�C���|�[�g������"
        Me.labelProgress.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'labelSubMsg
        '
        Me.labelSubMsg.Font = New System.Drawing.Font("�l�r �S�V�b�N", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.labelSubMsg.Location = New System.Drawing.Point(13, 32)
        Me.labelSubMsg.Name = "labelSubMsg"
        Me.labelSubMsg.Size = New System.Drawing.Size(408, 8)
        Me.labelSubMsg.TabIndex = 16
        Me.labelSubMsg.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'labelMainMsg
        '
        Me.labelMainMsg.Font = New System.Drawing.Font("�l�r �S�V�b�N", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.labelMainMsg.Location = New System.Drawing.Point(13, 8)
        Me.labelMainMsg.Name = "labelMainMsg"
        Me.labelMainMsg.Size = New System.Drawing.Size(408, 16)
        Me.labelMainMsg.TabIndex = 15
        Me.labelMainMsg.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'WaitDialog
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 12)
        Me.ClientSize = New System.Drawing.Size(434, 103)
        Me.ControlBox = False
        Me.Controls.Add(Me.progBarMeter)
        Me.Controls.Add(Me.labelProgress)
        Me.Controls.Add(Me.labelSubMsg)
        Me.Controls.Add(Me.labelMainMsg)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "WaitDialog"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "���s���E�E�E"
        Me.ResumeLayout(False)

    End Sub

#End Region
    ' ShowDialog���\�b�h�̃V���h�E�iWaitDialog�N���X�ł́AShowDialog���\�b�h�͎g�p�s�j
    Public Shadows Function ShowDialog() As DialogResult
        Debug.Assert(False, _
         "WaitDialog�N���X��ShowDialog���\�b�h�𗘗p�ł��܂���B" + vbCrLf + _
         "Show���\�b�h���g���ă��[�h���X�E�_�C�A���O���\�z���Ă��������B")
        Return DialogResult.Abort
    End Function

    ' Show���\�b�h�̃V���h�E
    Public Shadows Sub Show()
        ' �ϐ��̏�����
        Me.DialogResult = DialogResult.OK
        Me.bAborting = False

        MyBase.Show()
        Me.bShowing = True
    End Sub

    ' Close���\�b�h�̃V���h�E
    Public Shadows Sub Close()
        Me.bShowing = False
        MyBase.Close()
    End Sub

    ' �L�����Z���E�{�^���������ꂽ�Ƃ��̏���
    ' ������r���ŃL�����Z���i���f�j����B
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ' ���~����
        Abort()
    End Sub

    ' ���~�i�L�����Z���j����
    Private Sub Abort()
        Me.bAborting = True
        Me.DialogResult = DialogResult.Abort
    End Sub

    ' �_�C�A���O��������Ƃ��̏���
    ' �E��́m����n�{�^���������ꂽ�ꍇ�ɂ́A
    ' �m�L�����Z���n�{�^���Ɠ����悤�ɁA������r���ŃL�����Z���i���f�j����B
    Private Sub WaitDialog_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If bShowing = True Then
            ' �_�C�A���O�\�����Ȃ̂Œ��~�i�L�����Z���j���������s
            Abort()
            ' �܂��_�C�A���O�͕��Ȃ�
            e.Cancel = True
        Else
            ' �t�H�[���͕�����Ƃ���̂őf���ɕ���
            e.Cancel = False
        End If
    End Sub

    ' **************************************************************

    ' �������L�����Z���i���~�j����Ă��邩�ǂ����������l���擾����B
    ' �L�����Z�����ꂽ�ꍇ��true�B����ȊO��false�B
    Public ReadOnly Property IsAborting() As Boolean
        Get
            Return Me.bAborting
        End Get
    End Property

    ' ���C���E���b�Z�[�W�̃e�L�X�g��ݒ肷��B
    ' �����̊T�v��\������B
    ' �Ⴆ�΁A�t�@�C���̓]�����s���Ă���Ȃ�A�u�t�@�C����]�����Ă��܂��c�c�v�̂悤�ɕ\������B
    Public WriteOnly Property MainMsg() As String
        Set(ByVal Value As String)
            Me.labelMainMsg.Text = Value
        End Set
    End Property

    ' �T�u�E���b�Z�[�W�̃e�L�X�g��ݒ肷��B
    ' �ڍׂȏ������e��\������B
    ' �Ⴆ�΁A�t�@�C���]���Ȃ�A�]�����̌ʂ̃t�@�C�����i�ureadme.htm�v�ucontents.htm�v�Ȃǁj��\������B
    Public WriteOnly Property SubMsg() As String
        Set(ByVal Value As String)
            Me.labelSubMsg.Text = Value
        End Set
    End Property

    ' �i�s�󋵃��b�Z�[�W�̃e�L�X�g��ݒ肷��B
    ' �����̐i�s�󋵂Ƃ��āA�u�������̉������I������̂��v�u�S�̂̉������I������̂��v�Ȃǂ�\������B
    Public WriteOnly Property ProgressMsg() As String
        Set(ByVal Value As String)
            Me.labelProgress.Text = Value
        End Set
    End Property

    ' �i�s�󋵃��[�^�[�̌��݈ʒu��ݒ肷��B
    ' �Ⴆ�΁A������200�̍H�����������ꍇ�A���݂���200�̍H���̒��̂ǂ̈ʒu�ɂ��邩�������l�B
    ' ����l�́u0�v�B
    Public WriteOnly Property ProgressValue() As Integer
        Set(ByVal Value As Integer)
            Me.progBarMeter.Value = Value
        End Set
    End Property

    ' �i�s�󋵃��[�^�[�͈̔͂̍ő�l��ݒ肷��B
    ' ������200�̍H��������Ȃ�u200�v�ɂȂ�B
    ' ����l�́u100�v�B
    Public WriteOnly Property ProgressMax() As Integer
        Set(ByVal Value As Integer)
            Me.progBarMeter.Maximum = Value
        End Set
    End Property

    ' �i�s�󋵃��[�^�[�͈̔͂̍ŏ��l��ݒ肷��B
    ' ����l�́u0�v�B
    Public WriteOnly Property ProgressMin() As Integer
        Set(ByVal Value As Integer)
            Me.progBarMeter.Minimum = Value
        End Set
    End Property

    ' PerformStep���\�b�h���Ăяo�����Ƃ��ɁA�i�s�󋵃��[�^�[�̌��݈ʒu��i�߂�ʁiProgressValue�̑����l�j��ݒ肷��B
    ' �����H����200�ŁA5�̍H�����I������i�K�Ői�s�󋵃��[�^�[���X�V�������ꍇ�́u5�v�ɂ���B
    ' ����l�́u10�v�B
    Public WriteOnly Property ProgressStep() As Integer
        Set(ByVal Value As Integer)
            Me.progBarMeter.Step = Value
        End Set
    End Property

    ' �i�s�󋵃��[�^�[�̌��݈ʒu�iProgressValue�j��ProgressStep�v���p�e�B�̗ʂ����i�߂�B
    Public Sub PerformStep()
        Me.progBarMeter.PerformStep()
    End Sub
End Class
