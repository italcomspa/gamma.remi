VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLSMG_MANSCO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private FormSkProd                      As FRMMG_MANSCO

Implements Cinterface

' Variabili private di propietà
Private Gcls_VoceMenu                   As CLSFW_VOCE
Private Glng_Top                        As Long
Private Gbol_IsActive                   As Boolean
Private Glng_Left                       As Long
Private Gcon_Connect                    As ADODB.Connection
Private Gbol_CanInsert                  As Boolean
Private Gbol_CanDelete                  As Boolean
Private Gbol_CanUpdate                  As Boolean
Private Gbol_CanQbe                     As Boolean
Private Gint_WinState                   As Integer
Private Gbol_WinModal                   As Boolean
Private Gint_Mode                       As Integer
Private Gstr_MenuModulo                 As Variant
Private Gstr_MenuPadre                  As Variant
Private Gstr_MenuProg                   As Variant
Private Gstr_Caption                    As Variant
Private Glng_Parent                     As Long
Private Gcls_Global                     As CLSFW_Global
Private WithEvents ActiveNavigator      As CLSFW_Navigator
Attribute ActiveNavigator.VB_VarHelpID = -1
Private Gobj_ActiveDll                  As Cinterface
Private Gobj_StatusBar                  As Object
Private Gobj_ActiveFrame                As Object
Private Gbol_IsCalled                   As Boolean
Private Gvar_Filtro                     As Variant
Private Gstr_articolo                   As String
Private Gstr_variante                   As String

' Dichiarazione evento ritorno dati a programmi chiamanti
Public Event DataTransfer(ByVal Data As Variant)

'Passaggio parametri CODICE ARTICOLO
Public Property Let CodiceArticolo(RHS As String)
    On Error Resume Next
    
    Gstr_articolo = RHS
End Property
Public Property Get CodiceArticolo() As String
    On Error Resume Next
     
    CodiceArticolo = Gstr_articolo
End Property

'Passaggio parametri VARIANTE
Public Property Let Opzione(RHS As String)
    On Error Resume Next
    
    Gstr_variante = RHS
End Property
Public Property Get Opzione() As String
    On Error Resume Next
     
    Opzione = Gstr_variante
End Property

' Classe di interfaccia standard
Private Sub ActiveNavigator_BeforeExportToOffice(Cancel As Boolean, Pcls_ExpToOffice As Object)
    On Error Resume Next

    On Error Resume Next
    
    Dim Pvar_Campi          As Variant
    Dim Pvar_Dati           As Variant
    Dim Pint_Indice         As Integer
    Dim Pint_count          As Integer
    Dim I                   As Integer
    Dim Vis                 As Integer
    
    If Gobj_ActiveDll.IsActive = False Then
        Exit Sub
    End If
    
    If Not FormSkProd.Prst_Progressivi Is Nothing And FormSkProd.Prst_Progressivi.RecordCount > 0 Then
        '
        '   Scollego il datasource della griglia righe documenti
        '
        Set FormSkProd.GRID_GIACENZE.DataSource = Nothing
    
        Screen.MousePointer = vbHourglass
        FormSkProd.ActiveInterface.StatusBar.Panels(2).Text = "Traferimento dati in corso ..."
        
        For I = 0 To FormSkProd.GRID_GIACENZE.Columns.Count - 1
            If FormSkProd.GRID_GIACENZE.Columns.Item(I).Visible Then
                Vis = Vis + 1
            End If
        Next

        ReDim Pvar_Campi(Vis - 1)
        ReDim Pvar_Dati(0 To FormSkProd.Prst_Progressivi.RecordCount - 1, Vis - 1)
        
        Pint_count = 0
        For I = 0 To FormSkProd.GRID_GIACENZE.Columns.Count - 1
            If FormSkProd.GRID_GIACENZE.Columns.Item(I).Visible Then
                Pvar_Campi(Pint_count) = FormSkProd.GRID_GIACENZE.Columns.Item(I).Caption
                Pint_count = Pint_count + 1
            End If
        Next
    
        Pint_Indice = 0
        
        FormSkProd.Prst_Progressivi.MoveFirst

        While Not (FormSkProd.Prst_Progressivi.EOF)
            Pint_count = 0
            For I = 0 To FormSkProd.GRID_GIACENZE.Columns.Count - 1
                If FormSkProd.GRID_GIACENZE.Columns.Item(I).Visible Then
                    Pvar_Dati(Pint_Indice, Pint_count) = FormSkProd.Prst_Progressivi.Fields(FormSkProd.GRID_GIACENZE.Columns.Item(I).DataField).Value
                    Pint_count = Pint_count + 1
                End If
            Next
            Pint_Indice = Pint_Indice + 1
            FormSkProd.Prst_Progressivi.MoveNext
        Wend
        
        Pcls_ExpToOffice.AddMasterDetail Pvar_Campi, Pvar_Dati, "INQLIS. Articolo: " & FormSkProd.RTrimN(FormSkProd.TXT_CODART.Text)
        '
        '   Collego il datasource della griglia righe documenti
        '
        Set FormSkProd.GRID_GIACENZE.DataSource = FormSkProd.Prst_Progressivi
        
        Screen.MousePointer = vbNormal
        FormSkProd.ActiveInterface.StatusBar.Panels(2).Text = "Pronto"
    End If

'
'    If Gobj_ActiveDll.IsActive = False Then
'       Exit Sub
'    End If
'
'    Pcls_ExpToOffice.TitoloMaster = "Scheda Prodotti"
'    Pcls_ExpToOffice.AddMasterDetail Pvar_Campi, Pvar_Dati, "Interrogazione Partitari Lotti"
'
End Sub

Private Sub ActiveNavigator_BeforeChangeDitta(Cancel As Boolean)
    On Error Resume Next
    If Gobj_ActiveDll.IsActive = False Then
       Exit Sub
    End If
    MsgBox "Non è possibile cambiare ditta durante l'esecuzione del programma.", vbInformation, "Cambio Ditta"
    Cancel = True
    Err.Clear
End Sub

Private Sub ActiveNavigator_ButtonPressed(Cancel As Boolean, ButtonPress As FWBO_LibSrv.ButtonIndex)
On Error Resume Next
'MsgBox Gobj_ActiveDll.ActiveNavigator.ActiveInterface
'MsgBox FWBO_LibSrv.wsxMdiChild
'If Gobj_ActiveDll.ActiveNavigator.ActiveInterface = "Cinterface/CLSMG_INQLIS" Then
  
If Gobj_ActiveDll.ClsGlobal.Gcls_VoceMenu.Alias = "INQLIS" Then
  
  On Error Resume Next
  If FormSkProd.FME_CCS_SKPROD.Recordset.RecordCount = 1 Then
      Select Case ButtonPress
          Case tsNextRecord
              Call FormSkProd.ImpostaVirtualFrame(0)
          Case tsPreviousRecord
              Call FormSkProd.ImpostaVirtualFrame(1)
          Case tsFirstRecord
              Call FormSkProd.ImpostaVirtualFrame(2)
          Case tsLastRecord
              Call FormSkProd.ImpostaVirtualFrame(3)
      End Select
  End If
'Else
'    MsgBox "Non Entra"
'
End If

End Sub

Private Sub ActiveNavigator_InitializeLayoutObjectModel(Cancel As Boolean, ClsLayout As PRJFW_PRSLAYOUT.CLSFW_PRSLAYOUT)
On Error Resume Next

    If Not (Gobj_ActiveDll.IsActive) Then Exit Sub
    
    'CCS
    ClsLayout.AddStruttura "Form giacenze", FormSkProd, False, False, tsFormSelezione
    'CCS
    ClsLayout.AddStrutturaGrid "GIACENZE", FormSkProd.GRID_GIACENZE
End Sub

Private Sub ActiveNavigator_InitializeScriptObjectModel(Cancel As Boolean, ClsScript As FWUO_TMSDEVELOP.CLSFW_PRSVBSCRIPT___v2)
On Error Resume Next

    If Not (Gobj_ActiveDll.IsActive) Then Exit Sub
    
    Call ClsScript.AddObject(FormSkProd.Name, FormSkProd, 0, "Scheda Prodotti")
    Gobj_ActiveDll.ActiveFrame.Name = "DTS Scheda Prodotti"
    Call ClsScript.AddObject(Gobj_ActiveDll.ActiveFrame.Name, Gobj_ActiveDll.ActiveFrame, FWUO_TMSDEVELOP.tsDataSource, Gobj_ActiveDll.ActiveFrame.Name)
End Sub

Private Property Set Cinterface_ActiveDll(RHS As FWBO_LibSrv.Cinterface)
    On Error Resume Next
     
    Set Gobj_ActiveDll = RHS
End Property
Private Property Get Cinterface_ActiveDll() As FWBO_LibSrv.Cinterface
    On Error Resume Next
     
    Set Cinterface_ActiveDll = Gobj_ActiveDll
End Property

Private Property Get Cinterface_ActiveFrame() As Object
    On Error Resume Next
     
    Set Cinterface_ActiveFrame = Gobj_ActiveFrame
End Property
Private Property Set Cinterface_ActiveFrame(RHS As Object)
    On Error Resume Next
     
    Set Gobj_ActiveFrame = RHS
End Property

Private Property Set Cinterface_ActiveNavigator(RHS As FWBO_LibSrv.CLSFW_Navigator)
    On Error Resume Next
    
    Set ActiveNavigator = RHS
End Property

Private Property Get Cinterface_ActiveNavigator() As FWBO_LibSrv.CLSFW_Navigator
    On Error Resume Next
    
    Set Cinterface_ActiveNavigator = ActiveNavigator
End Property

Private Property Let Cinterface_CanDelete(ByVal RHS As Boolean)
    On Error Resume Next
    
    Gbol_CanDelete = RHS
End Property
Private Property Get Cinterface_CanDelete() As Boolean
     On Error Resume Next
     
     Cinterface_CanDelete = Gbol_CanDelete
End Property

Private Property Let Cinterface_CanInsert(ByVal RHS As Boolean)
    On Error Resume Next
    
    Gbol_CanInsert = RHS
End Property
Private Property Get Cinterface_CanInsert() As Boolean
    On Error Resume Next
    
    Cinterface_CanInsert = Gbol_CanInsert
End Property

Private Property Let Cinterface_CanQbe(ByVal RHS As Boolean)
    On Error Resume Next
    
    Gbol_CanQbe = RHS
End Property
Private Property Get Cinterface_CanQbe() As Boolean
    On Error Resume Next
    
    Cinterface_CanQbe = Gbol_CanQbe
End Property

Private Property Let Cinterface_CanUpdate(ByVal RHS As Boolean)
    On Error Resume Next
    
    Gbol_CanUpdate = RHS
End Property

Private Property Get Cinterface_CanUpdate() As Boolean
    On Error Resume Next
    
    Cinterface_CanUpdate = Gbol_CanUpdate
End Property

Private Property Let Cinterface_Caption(ByVal RHS As Variant)
    On Error Resume Next
    
    Gstr_Caption = RHS
End Property

Private Property Get Cinterface_Caption() As Variant
    On Error Resume Next
    
    Cinterface_Caption = Gstr_Caption
End Property

Private Function Cinterface_CloseForm() As Variant
    
    ' Metodo chiusira form
    Cinterface_CloseForm = False
    Unload FormSkProd
    If FormSkProd.ActiveInterface Is Nothing Then
      Cinterface_CloseForm = True
    End If
End Function

Private Property Get Cinterface_ClsGlobal() As FWBO_LibSrv.CLSFW_Global
    On Error Resume Next
    
    Set Cinterface_ClsGlobal = Gcls_Global
End Property

Private Property Set Cinterface_ClsGlobal(RHS As FWBO_LibSrv.CLSFW_Global)
    On Error Resume Next
    
    Set Gcls_Global = RHS
End Property

Private Property Set Cinterface_ClsVoceMenu(RHS As FWBO_GESMODULI.CLSFW_VOCE)
    On Error Resume Next
    
    Set Gcls_VoceMenu = RHS
End Property

Private Property Get Cinterface_ClsVoceMenu() As FWBO_GESMODULI.CLSFW_VOCE
    On Error Resume Next
    
    Set Cinterface_ClsVoceMenu = Gcls_VoceMenu
End Property

Private Property Get Cinterface_Connection() As ADODB.Connection
    On Error Resume Next
    
    Set Cinterface_Connection = Gcon_Connect
End Property

Private Property Set Cinterface_Connection(RHS As ADODB.Connection)
    On Error Resume Next
    
    Set Gcon_Connect = RHS
End Property

Private Property Get Cinterface_Cwindow() As FWBO_LibSrv.CLSFW_Window

End Property

Private Property Let Cinterface_Filtro(ByVal RHS As Variant)
    On Error Resume Next
    
    Gvar_Filtro = RHS
End Property

Private Property Get Cinterface_Filtro() As Variant
    On Error Resume Next
    
    Set Cinterface_Filtro = Gvar_Filtro
End Property

Private Property Let Cinterface_hwndParent(ByVal RHS As Long)
    On Error Resume Next
    
    Glng_Parent = RHS
End Property
Private Property Get Cinterface_hwndParent() As Long
    On Error Resume Next
    
    Cinterface_hwndParent = Glng_Parent
End Property

Private Property Let Cinterface_IsActive(ByVal RHS As Boolean)
    On Error Resume Next
    
    Gbol_IsActive = RHS
End Property

Private Property Get Cinterface_IsActive() As Boolean
    On Error Resume Next
    
    Cinterface_IsActive = Gbol_IsActive
End Property

Private Property Let Cinterface_IsCalled(ByVal RHS As Boolean)
    On Error Resume Next
    
    Gbol_IsCalled = RHS
End Property

Private Property Get Cinterface_IsCalled() As Boolean
    On Error Resume Next
    
    Cinterface_IsCalled = Gbol_IsCalled
End Property

Private Property Let Cinterface_MenuModulo(ByVal RHS As Variant)
    On Error Resume Next
    
    Gstr_MenuModulo = RHS
End Property

Private Property Get Cinterface_MenuModulo() As Variant
    On Error Resume Next
    
    Cinterface_MenuModulo = Gstr_MenuModulo
End Property

Private Property Let Cinterface_MenuPadre(ByVal RHS As Variant)
    On Error Resume Next
    
    Gstr_MenuPadre = RHS
End Property

Private Property Get Cinterface_MenuPadre() As Variant
    On Error Resume Next
    
    Cinterface_MenuPadre = Gstr_MenuPadre
End Property

Private Property Let Cinterface_MenuProg(ByVal RHS As Variant)
    On Error Resume Next
    
    Gstr_MenuProg = RHS
End Property

Private Property Get Cinterface_MenuProg() As Variant
    On Error Resume Next
    
    Cinterface_MenuProg = Gstr_MenuProg
End Property

Private Property Let Cinterface_ProgramMode(ByVal RHS As FWBO_LibSrv.Mode)
    On Error Resume Next
    
    Gint_Mode = RHS
End Property

Private Property Get Cinterface_ProgramMode() As FWBO_LibSrv.Mode
    On Error Resume Next
    
    Cinterface_ProgramMode = Gint_Mode
End Property

Private Function Cinterface_ShowForm() As Variant
    On Error Resume Next
    
    ' Impostazione modalità di apertura (Max,Min e normal)
    If FormSkProd Is Nothing Then
        Set FormSkProd = New FRMMG_MANSCO
        Set FormSkProd.ActiveInterface = Gobj_ActiveDll
        Set FormSkProd.ActiveClass = Gobj_ActiveDll
        FormSkProd.WindowState = Gint_WinState
        
        ' Impostazione modalità di apertura programma (Modal/Modless)
        ' Per la modalità Modless occorre non agganciare la finestra al
        ' padre altrimenti resta bloccata !!!!!!!!
        FormSkProd.Caption = Gstr_Caption
        If Gbol_WinModal = False Then
             FormSkProd.Show 0
        Else
             FormSkProd.Show 1
        End If
    Else
        FormSkProd.Show
    End If

End Function

Private Property Set Cinterface_StatusBar(RHS As Object)
    On Error Resume Next
    
    Set Gobj_StatusBar = RHS
End Property

Private Property Get Cinterface_StatusBar() As Object
    On Error Resume Next
    
    Set Cinterface_StatusBar = Gobj_StatusBar
End Property

' Metodo che consente di alzare l'evento di trasferimento dati
Private Function Cinterface_TransferData(ByVal Data As Variant) As Variant
    On Error Resume Next
    
    RaiseEvent DataTransfer(Data)
End Function

Private Property Get Cinterface_WindowModal() As Boolean
    On Error Resume Next
    
    Cinterface_WindowModal = Gbol_WinModal
End Property

Private Property Let Cinterface_WindowModal(ByVal RHS As Boolean)
    On Error Resume Next
    
    Gbol_WinModal = RHS
End Property

Private Property Let Cinterface_WindowState(ByVal RHS As FWBO_LibSrv.WindowState)
    On Error Resume Next
    
    Gint_WinState = RHS
End Property

Private Property Get Cinterface_WindowState() As FWBO_LibSrv.WindowState
    On Error Resume Next
    
    Cinterface_WindowState = Gint_WinState
End Property

Private Sub Class_Initialize()
    On Error Resume Next
    
    Set Gobj_ActiveDll = Me
End Sub

Private Property Let Cinterface_Top(ByVal RHS As Variant)
    On Error Resume Next
    
    Glng_Top = RHS
End Property

Private Property Get Cinterface_Top() As Variant
    On Error Resume Next
    
    Cinterface_Top = Glng_Top
End Property

Private Property Let Cinterface_Left(ByVal RHS As Variant)
    On Error Resume Next
    
    Glng_Left = RHS
End Property

Private Property Get Cinterface_Left() As Variant
    On Error Resume Next
    
    Cinterface_Left = Glng_Left
End Property

Private Sub Class_Terminate()
On Error Resume Next
    Set Gcls_VoceMenu = Nothing
    Set Gcls_Global = Nothing
    Set ActiveNavigator = Nothing
    Set Gobj_ActiveDll = Nothing
    Set FormSkProd = Nothing
End Sub
