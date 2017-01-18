Attribute VB_Name = "MenuCreator"
Option Explicit

Public Const NAME_MENU As String = "�����������"
Public Const NAME_MINPART As String = "���������� ����������� ������"
Public Const NAME_MINPART_UNLOAD_DZ As String = "��������� ������ �� ��"
Public Const NAME_MINPART_UNLOAD_ZDZ As String = "��������� ������ �� ���"
Public Const NAME_MINPART_LOAD_CHANGE As String = "��������� ���������"
Public Const NAME_MINPART_UNLOAD_POSITION As String = "��������� ������ �� �������"
Public Const NAME_MINPART_LOAD_TEMPLATE As String = "��������� ������"
Public Const NAME_MINPART_UNLOAD_ALL As String = "��������� ��� ������"
'---------------------------------------------------------------------------------------
Public Const NAME_CONTROL_NOVELTY As String = "�������� ����� �������"
Public Const NAME_NOVELTY_UNLOAD_DZ As String = "��������� ������ �� ��"
Public Const NAME_NOVELTY_UNLOAD_ZDZ As String = "��������� ������ �� ���"
Public Const NAME_NOVELTY_LOAD_CHANGE As String = "��������� ���������"
Public Const NAME_NOVELTY_UNLOAD_POSITION As String = "��������� ������ �� ��������"
Public Const NAME_NOVELTY_LOAD_TEMPLATE As String = "��������� ������"
Public Const NAME_NOVELTY_UNLOAD_ALL As String = "��������� ��� ������"
'---------------------------------------------------------------------------------------
Public Const NAME_CONTROL_SHOWBOX As String = "������� ������� � ��������"
Public Const NAME_SHOWBOX_UNLOAD_DZ As String = "��������� ������ �� ��"
Public Const NAME_SHOWBOX_UNLOAD_ZDZ As String = "��������� ������ �� ���"
Public Const NAME_SHOWBOX_LOAD_CHANGE As String = "��������� ���������"
Public Const NAME_SHOWBOX_UNLOAD_POSITION As String = "��������� ������ �� ��������"
Public Const NAME_SHOWBOX_LOAD_TEMPLATE As String = "��������� ������"
Public Const NAME_SHOWBOX_UNLOAD_ALL As String = "��������� ��� ������"

Public Sub MenuBuild_dz()

    Dim bar As CommandBar
    Dim barMenu As CommandBarControl
    
    '�������� ����������������� ���� ���� ��� ��� ����������
    Call MenuKill_dz
    '�������� ������ ����, ���������� ���������� ����
    Set bar = Application.CommandBars("Worksheet Menu Bar")
    ' �������� ����������� ����
    Set barMenu = bar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    
    Dim version As String
    version = getVersion()
    
    With barMenu
        .Caption = NAME_MENU & version
    End With
    '���� ���� �����������, �� ����� �� ���������
    If barMenu Is Nothing Then Exit Sub
    
    Call createMinPart(barMenu)
    Call createControlNovelty(barMenu)
    Call createShowBox(barMenu)
    
    Set barMenu = Nothing
    Set bar = Nothing
End Sub

Sub MenuKill_dz()
    On Error Resume Next
    '��� ��� �� ��������� ������ � ��������, ������� ��� ��� �������� ���� ��������
    Dim pos As Integer
    Dim c
    For Each c In Application.CommandBars("Worksheet Menu Bar").Controls
        pos = InStr(c.Caption, NAME_MENU)
        If pos <> 0 Then
            c.Delete
        End If
    Next c
    On Error GoTo 0
End Sub


Sub createMinPart(barMenu As CommandBarControl)
    '������� ���� ���������
    Dim minPart As CommandBarControl
    Set minPart = barMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    With minPart
        .Caption = NAME_MINPART
    End With
    
    '���������� ������ ����
    With minPart.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = NAME_MINPART_UNLOAD_DZ
        .OnAction = ThisWorkbook.FullName & "!ShowFrames.showFrameUnloadDZMinpart"
    End With
    
    With minPart.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = NAME_MINPART_UNLOAD_ZDZ
        .OnAction = ThisWorkbook.FullName & "!ShowFrames.showFrameUnloadZDZMinpart"
    End With
    
    With minPart.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = NAME_MINPART_UNLOAD_POSITION
        .OnAction = ThisWorkbook.FullName & "!UnloadMP.unloadPosition"
    End With
       
    With minPart.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = NAME_MINPART_LOAD_CHANGE
        .OnAction = ThisWorkbook.FullName & "!LoadMP.loadChange"
        .BeginGroup = True
    End With
    
    Dim b As Boolean: b = GrantHolder.nowUserIsAdminMinPart
    If (b) Then
        With minPart.Controls.Add(Type:=msoControlButton, Temporary:=True)
            .Caption = NAME_MINPART_LOAD_TEMPLATE
            .OnAction = ThisWorkbook.FullName & "!LoaderTemplate.loadMPFile"
            .BeginGroup = True
        End With
        With minPart.Controls.Add(Type:=msoControlButton, Temporary:=True)
            .Caption = NAME_MINPART_UNLOAD_ALL
            .OnAction = ThisWorkbook.FullName & "!UnloadMP.unloadAll"
        End With
    End If
    
End Sub

Sub createControlNovelty(barMenu As CommandBarControl)
    Dim controlNovelty As CommandBarControl
    Set controlNovelty = barMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    With controlNovelty
        .Caption = NAME_CONTROL_NOVELTY
    End With
    
    '���������� ������ ����
    With controlNovelty.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = NAME_NOVELTY_UNLOAD_DZ
        .OnAction = ThisWorkbook.FullName & "!ShowFrames.showFrameUnloadDZControlNovelty"
    End With
    
    With controlNovelty.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = NAME_NOVELTY_UNLOAD_ZDZ
        .OnAction = ThisWorkbook.FullName & "!ShowFrames.showFrameUnloadZDZControlNovelty"
    End With
    
   
    With controlNovelty.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = NAME_NOVELTY_UNLOAD_POSITION
        .OnAction = ThisWorkbook.FullName & "!UnloadCN.unloadPosition"
    End With
    
    With controlNovelty.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = NAME_NOVELTY_LOAD_CHANGE
        .OnAction = ThisWorkbook.FullName & "!LoadCN.loadChange"
        .BeginGroup = True
    End With
    
    Dim b As Boolean: b = GrantHolder.nowUserIsAdminControlNovelty
    If (b) Then
        With controlNovelty.Controls.Add(Type:=msoControlButton, Temporary:=True)
            .Caption = NAME_NOVELTY_LOAD_TEMPLATE
            .OnAction = ThisWorkbook.FullName & "!LoaderTemplate.loadCNFile"
            .BeginGroup = True
        End With
        With controlNovelty.Controls.Add(Type:=msoControlButton, Temporary:=True)
            .Caption = NAME_NOVELTY_UNLOAD_ALL
            .OnAction = ThisWorkbook.FullName & "!UnloadCN.unloadAll"
        End With
    End If
End Sub

Sub createShowBox(barMenu As CommandBarControl)
    Dim showBox As CommandBarControl
    Set showBox = barMenu.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    With showBox
        .Caption = NAME_CONTROL_SHOWBOX
    End With
    
    '���������� ������ ����
    With showBox.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = NAME_SHOWBOX_UNLOAD_DZ
        .OnAction = ThisWorkbook.FullName & "!ShowFrames.showFrameUnloadDZShowBox"
    End With
    
    With showBox.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = NAME_SHOWBOX_UNLOAD_ZDZ
        .OnAction = ThisWorkbook.FullName & "!ShowFrames.showFrameUnloadZDZShowBox"
    End With
    
   
    With showBox.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = NAME_SHOWBOX_UNLOAD_POSITION
        .OnAction = ThisWorkbook.FullName & "!UnloadShowBox.unloadPosition"
    End With
    
    With showBox.Controls.Add(Type:=msoControlButton, Temporary:=True)
        .Caption = NAME_SHOWBOX_LOAD_CHANGE
        .OnAction = ThisWorkbook.FullName & "!LoadShowBox.loadChange"
        .BeginGroup = True
    End With
    
    Dim b As Boolean: b = GrantHolder.nowUserIsAdminShowBox
    If (b) Then
        With showBox.Controls.Add(Type:=msoControlButton, Temporary:=True)
            .Caption = NAME_SHOWBOX_LOAD_TEMPLATE
            .OnAction = ThisWorkbook.FullName & "!LoaderTemplate.loadShowBoxExecute"
            .BeginGroup = True
        End With
        With showBox.Controls.Add(Type:=msoControlButton, Temporary:=True)
            .Caption = NAME_SHOWBOX_UNLOAD_ALL
            .OnAction = ThisWorkbook.FullName & "!UnloadShowBox.unloadAll"
        End With
    End If
End Sub

Private Function getVersion() As String
    Dim nameBook As String
    nameBook = ThisWorkbook.name
    Dim position As Integer
    Dim version As String
    
    position = InStr(1, nameBook, "_")
    If position = 0 Then
        getVersion = ""
        Exit Function
    Else
        version = Mid(nameBook, position, Len(nameBook) - position - 4) '5 ��� ���������� .xlam
        getVersion = version
        Exit Function
    End If
End Function
