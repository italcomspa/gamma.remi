VERSION 5.00
Object = "{0EF4EAA6-2617-11D2-A1C0-0060082875F9}#4.12#0"; "TMS_COMBOBOX.ocx"
Object = "{0EF4EA3A-2617-11D2-A1C0-0060082875F9}#8.11#0"; "TMS_EDIT.ocx"
Object = "{5032AB27-52C8-11D2-A1C0-0060082875F9}#4.16#0"; "TMS_EDITM.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{589B69E4-BA68-11D1-9B0E-006097A80EFD}#5.5#0"; "tms_groupbox.ocx"
Object = "{F53BE214-7AC6-11D0-9B0E-006097A80EFD}#6.10#0"; "TMS_LABEL.ocx"
Object = "{0EF4E915-2617-11D2-A1C0-0060082875F9}#7.19#0"; "TMS_RICHTEXTBOX.ocx"
Object = "{31930FDA-530C-11D2-A1C0-0060082875F9}#2.32#0"; "TMS_ARTICOLO.ocx"
Object = "{52AC1257-7978-11D2-A807-006097A80EFD}#2.30#0"; "TMS_EDITVARIANTE.ocx"
Object = "{CBAF6F53-3C3D-11D4-AA70-000629C16DEA}#2.4#0"; "MDIActiveXS.ocx"
Object = "{B473387D-A75F-4A83-9879-4A8FE48EE80F}#1.8#0"; "TMS_TBARMENU.ocx"
Begin VB.Form FRMMG_SITGIACENZE 
   Caption         =   "INQLIS"
   ClientHeight    =   6645
   ClientLeft      =   -3960
   ClientTop       =   345
   ClientWidth     =   11625
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMMG_SITGIACENZE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   11625
   StartUpPosition =   3  'Windows Default
   Begin PRJFW_EDITM.TXT_EDITM TXT_CODARTFOR 
      Height          =   300
      Left            =   1320
      TabIndex        =   48
      Top             =   1710
      Width           =   3720
      _ExtentX        =   6535
      _ExtentY        =   529
      IsLookup        =   -1  'True
      DisplayFormat   =   "Maiuscolo"
      MaxChar         =   25
      Numerico        =   0   'False
      Carattere       =   0   'False
      IsDbField       =   0   'False
      NumRighe        =   0
      MaxWidth        =   25
   End
   Begin PRJFW_TBARMENU.TMS_TBARMENU CMD_NUOVO 
      Height          =   345
      Left            =   10530
      TabIndex        =   3
      Top             =   6300
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      Caption         =   "&Nuovo"
      IsMenuPopup     =   0   'False
   End
   Begin MDIinActiveX.MDIActiveX MDIActiveX1 
      Left            =   2190
      Top             =   6330
      _ExtentX        =   847
      _ExtentY        =   794
      BorderStyle     =   0
   End
   Begin PRJFW_RICHTEXTBOX.TmsRichTextBox TXT_DESCART 
      Height          =   300
      Left            =   4650
      TabIndex        =   4
      ToolTipText     =   "Descrizione articolo"
      Top             =   150
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   529
      Enabled         =   0   'False
      MaxChar         =   72
      Numerico        =   0   'False
      Carattere       =   0   'False
      IsDbField       =   0   'False
      Caption         =   "Descrizione articolo"
      Object.Tag             =   "Descrizione articolo"
      MaxWidth        =   55
   End
   Begin MSDataGridLib.DataGrid GRID_GIACENZE 
      Height          =   1515
      Left            =   0
      TabIndex        =   13
      Top             =   2130
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   2672
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   52
      BeginProperty Column00 
         DataField       =   "DITTA"
         Caption         =   "Ditta"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "VARIANTE"
         Caption         =   "Variante"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "COD_PROGETTO"
         Caption         =   "Progetto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "DESCR_PROGETTO"
         Caption         =   "Descrizione progetto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "DEPOSITO"
         Caption         =   "Dep"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "DESCR_DEPOSITO"
         Caption         =   "Descrizione deposito"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "QGIACATT"
         Caption         =   "Giac.attuale"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   """€"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "QGIACEFF"
         Caption         =   "Giac.effettiva"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "QORDFOR"
         Caption         =   "Ord.fornitore"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "QIMPCLI"
         Caption         =   "Imp.cliente"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "QDISPONIB"
         Caption         =   "Disponibile"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "QGIACINI"
         Caption         =   "Giac.iniziale"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "QGIACFIS"
         Caption         =   "Giac.fiscale"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "QORDPROD"
         Caption         =   "Ord. Prod."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "QIMPPROD"
         Caption         =   "Imp. Prod."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "QIMPCLAVFOR"
         Caption         =   "Imp. C.Lav For"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "QPREIMPFOR"
         Caption         =   "Preimp.fo.pr."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column17 
         DataField       =   "QPREIMPCLI"
         Caption         =   "Preimp.cliente"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column18 
         DataField       =   "QBLOCSPED"
         Caption         =   "Blocco.sped."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column19 
         DataField       =   "QDACONTR"
         Caption         =   "Da controllare"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column20 
         DataField       =   "QDAVAL"
         Caption         =   "Da valorizzare"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column21 
         DataField       =   "QENTCVIS"
         Caption         =   "Entr.c/visione"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column22 
         DataField       =   "QENTCRIP"
         Caption         =   "Entr.c/ripar."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column23 
         DataField       =   "QENCDEP"
         Caption         =   "Entr.c/dep."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column24 
         DataField       =   "QENCNOLO"
         Caption         =   "Entr.c/nolo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column25 
         DataField       =   "QUSCCVIS"
         Caption         =   "Usc.c/visione"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column26 
         DataField       =   "QUSCCRIP"
         Caption         =   "Usc.c/rip."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column27 
         DataField       =   "QUSCDEP"
         Caption         =   "Usc.c/dep."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column28 
         DataField       =   "QUSCNOLO"
         Caption         =   "Usc.c/nolo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column29 
         DataField       =   "QCARACQ"
         Caption         =   "Car.acq."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column30 
         DataField       =   "QCARESORCLI"
         Caption         =   "Car.r.cli."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column31 
         DataField       =   "QCARPROD"
         Caption         =   "Car.prod."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column32 
         DataField       =   "QCARCLAVCLI"
         Caption         =   "Car.c/lav.cl."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column33 
         DataField       =   "QCARCLAVFOR"
         Caption         =   "Car.c/lav.fo."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column34 
         DataField       =   "QCAROMAG"
         Caption         =   "Car.omaggio"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column35 
         DataField       =   "QCARGENER"
         Caption         =   "Car.gener."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column36 
         DataField       =   "QCARTRASF"
         Caption         =   "Car.tr.dep."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column37 
         DataField       =   "QCARSOST"
         Caption         =   "Car.r.sost."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column38 
         DataField       =   "QCARLIB1"
         Caption         =   "Car. lib.1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column39 
         DataField       =   "QCARLIB2"
         Caption         =   "Car. lib.2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column40 
         DataField       =   "QSCAVEN"
         Caption         =   "Scar.vend."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column41 
         DataField       =   "QSCASCART"
         Caption         =   "Scar.scar."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column42 
         DataField       =   "QSCAOMAGQ"
         Caption         =   "Scar.omaggio"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column43 
         DataField       =   "QSCACLAVCLI"
         Caption         =   "Scar.c/lav.cl."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column44 
         DataField       =   "QSCACLAVFOR"
         Caption         =   "Scar.c/lav.fo."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column45 
         DataField       =   "QSCAPROD"
         Caption         =   "Scar.prod."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column46 
         DataField       =   "QSCARESOFOR"
         Caption         =   "Scar.r.for."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column47 
         DataField       =   "QSCAGENER"
         Caption         =   "Scar.gener."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column48 
         DataField       =   "QSCATRASF"
         Caption         =   "Scar.tr.dep."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column49 
         DataField       =   "QSCASOST"
         Caption         =   "Scar.sost."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column50 
         DataField       =   "QSCALIB1"
         Caption         =   "Scar.lib.1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column51 
         DataField       =   "QSCALIB2"
         Caption         =   "Scar.lib.2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   3
         Locked          =   -1  'True
         BeginProperty Column00 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   14,74
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   345,26
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1604,976
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column11 
         EndProperty
         BeginProperty Column12 
         EndProperty
         BeginProperty Column13 
         EndProperty
         BeginProperty Column14 
         EndProperty
         BeginProperty Column15 
         EndProperty
         BeginProperty Column16 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column17 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column18 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column19 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column20 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column21 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column22 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column23 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column24 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column25 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column26 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column27 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column28 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column29 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column30 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column31 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column32 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column33 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column34 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column35 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column36 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column37 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column38 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column39 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column40 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column41 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column42 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column43 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column44 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column45 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column46 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column47 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column48 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column49 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column50 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column51 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1154,835
         EndProperty
      EndProperty
   End
   Begin PRJFW_ARTICOLO.TxtArticolo TXT_CODART 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "Codice articolo"
      Top             =   150
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   529
      RunMenuEntry    =   -1  'True
      MaxChar         =   25
      Obbligatorio    =   -1  'True
      Numerico        =   0   'False
      Carattere       =   0   'False
      DBField         =   "MG66_CODART"
      IsQbe           =   -1  'True
      IsDecode        =   -1  'True
      Caption         =   "Codice articolo"
      Object.Tag             =   "Codice articolo"
      MaxWidth        =   15
      CanReturnRecordSet=   -1  'True
   End
   Begin PRJFW_TBARMENU.TMS_TBARMENU CMD_ELABORA 
      Height          =   345
      Left            =   9420
      TabIndex        =   2
      Top             =   6300
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      Caption         =   "&Elabora"
      IsMenuPopup     =   0   'False
   End
   Begin PRJFW_EDITVARIANTE.TXT_EDITVARIANTE TXT_OPZIONE 
      Height          =   300
      Left            =   12060
      TabIndex        =   18
      ToolTipText     =   "Variante articolo"
      Top             =   1290
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   529
      Enabled         =   0   'False
      Object.Visible         =   0   'False
      DBField         =   "VARIANTE"
      Caption         =   "Variante articolo"
      TipoVariante    =   0
   End
   Begin PRJFW_RICHTEXTBOX.TmsRichTextBox TXT_DESCARTEST 
      Height          =   300
      Left            =   4650
      TabIndex        =   20
      ToolTipText     =   "Descrizione articolo estesa"
      Top             =   480
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   529
      Enabled         =   0   'False
      MaxChar         =   72
      Numerico        =   0   'False
      Carattere       =   0   'False
      IsDbField       =   0   'False
      Caption         =   "Descrizione articolo estesa"
      NumRighe        =   10
      IsExpand        =   -1  'True
      Object.Tag             =   "Descrizione articolo estesa"
      MaxWidth        =   55
   End
   Begin PRJFW_ARTICOLO.TxtArticolo TXT_CODARTSOST 
      Height          =   300
      Left            =   12180
      TabIndex        =   26
      ToolTipText     =   "Codice articolo"
      Top             =   4350
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   529
      RunMenuEntry    =   -1  'True
      Enabled         =   0   'False
      MaxChar         =   25
      Object.Visible         =   0   'False
      Numerico        =   0   'False
      Carattere       =   0   'False
      IsDbField       =   0   'False
      IsQbe           =   -1  'True
      IsDecode        =   -1  'True
      Caption         =   "Codice articolo"
      Object.Tag             =   "Codice articolo"
      MaxWidth        =   25
      CanReturnRecordSet=   -1  'True
   End
   Begin PRJFW_RICHTEXTBOX.TmsRichTextBox TXT_DESCARTCF 
      Height          =   300
      Left            =   12030
      TabIndex        =   27
      ToolTipText     =   "Descrizione articolo"
      Top             =   1410
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
      Enabled         =   0   'False
      MaxChar         =   72
      Object.Visible         =   0   'False
      Numerico        =   0   'False
      Carattere       =   0   'False
      IsDbField       =   0   'False
      Caption         =   "Descrizione articolo"
      Object.Tag             =   "Descrizione articolo"
      MaxWidth        =   5
   End
   Begin PRJFW_RICHTEXTBOX.TmsRichTextBox TXT_CODARTSOSTDES 
      Height          =   300
      Left            =   12270
      TabIndex        =   28
      ToolTipText     =   "Descrizione articolo"
      Top             =   2340
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
      Enabled         =   0   'False
      MaxChar         =   72
      Object.Visible         =   0   'False
      Numerico        =   0   'False
      Carattere       =   0   'False
      IsDbField       =   0   'False
      Caption         =   "Descrizione articolo"
      Object.Tag             =   "Descrizione articolo"
      MaxWidth        =   5
   End
   Begin MSDataGridLib.DataGrid GRID_LISVEN 
      Height          =   1005
      Left            =   1650
      TabIndex        =   29
      Top             =   3660
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1773
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "LI10_NUMLIST"
         Caption         =   "Num. Lis."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "LI10_PREZZO"
         Caption         =   "Prezzo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "LI10_DATAINIZIOVAL"
         Caption         =   "Data Ini."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "LI10_DATAFINEVAL"
         Caption         =   "Data Fine"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   3
         Locked          =   -1  'True
         Size            =   242
         BeginProperty Column00 
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1184,882
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid GRID_LISACQ 
      Height          =   795
      Left            =   1650
      TabIndex        =   42
      Top             =   5490
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   1402
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "LI11_PROG"
         Caption         =   "Prog."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "LI11_DATAREG"
         Caption         =   "Data Reg"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "LI11_CODICE_CG08"
         Caption         =   "Valuta"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "LI11_PREZZO"
         Caption         =   "Prezzo Acq"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "LI11_SC1PER"
         Caption         =   "Sc1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "LI11_SC2PER"
         Caption         =   "Sc2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "LI11_SC3PER"
         Caption         =   "Sc3"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "LI11_SC4PER"
         Caption         =   "Sc4"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "LI11_SCIMP"
         Caption         =   "Sc. Imp"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "LI11_MAG1PER"
         Caption         =   "Mag %"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "LI11_MAGIMP"
         Caption         =   "Mag. Imp"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "PREZZO_NETTO"
         Caption         =   "Costo Netto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   2
         Locked          =   -1  'True
         Size            =   242
         BeginProperty Column00 
            ColumnWidth     =   780,095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   959,811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   450,142
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   464,882
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   464,882
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   480,189
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   884,976
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   569,764
         EndProperty
         BeginProperty Column10 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   989,858
         EndProperty
         BeginProperty Column11 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1124,787
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid GRID_LISACQ_TOT 
      Height          =   795
      Left            =   1650
      TabIndex        =   43
      Top             =   4680
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   1402
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "LI10_PREZZO"
         Caption         =   "Prezzo Acq"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "LI10_SC1PER"
         Caption         =   "Sc1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "LI10_SC2PER"
         Caption         =   "Sc2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "LI10_SC3PER"
         Caption         =   "Sc3"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "LI10_SC4PER"
         Caption         =   "Sc4"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "LI10_SCIMP"
         Caption         =   "Sc. Imp"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "LI10_MAG1PER"
         Caption         =   "Mag %"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "LI10_MAGIMP"
         Caption         =   "Mag. Imp"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "PREZZO_NETTO"
         Caption         =   "Costo Netto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   2
         Locked          =   -1  'True
         Size            =   242
         BeginProperty Column00 
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   555,024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   569,764
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   555,024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   569,764
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   959,811
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   675,213
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1124,787
         EndProperty
      EndProperty
   End
   Begin PRJFW_ARTICOLO.TxtArticolo TxtArticolo1 
      Height          =   300
      Left            =   12120
      TabIndex        =   45
      ToolTipText     =   "Codice articolo"
      Top             =   4800
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   529
      RunMenuEntry    =   -1  'True
      Enabled         =   0   'False
      MaxChar         =   25
      Object.Visible         =   0   'False
      Numerico        =   0   'False
      Carattere       =   0   'False
      IsDbField       =   0   'False
      IsQbe           =   -1  'True
      IsDecode        =   -1  'True
      Caption         =   "Codice articolo"
      Object.Tag             =   "Codice articolo"
      MaxWidth        =   25
      CanReturnRecordSet=   -1  'True
   End
   Begin MSDataGridLib.DataGrid GRID_ARTSOST 
      Height          =   1005
      Left            =   8490
      TabIndex        =   46
      Top             =   1440
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   1773
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "MG85_CODARTSOST_MG66"
         Caption         =   "Articolo Sostitutivo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "MG85_DATASOST"
         Caption         =   "Data"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   3
         Locked          =   -1  'True
         Size            =   242
         BeginProperty Column00 
            ColumnWidth     =   2115,213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   945,071
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid GRID_ARTALT 
      Height          =   1005
      Left            =   8490
      TabIndex        =   47
      Top             =   2580
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   1773
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   8388608
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "MG84_CODARTALT_MG66"
         Caption         =   "Articolo Alternativo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "MG84_CODRAGALT"
         Caption         =   "Prog"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   3
         Locked          =   -1  'True
         Size            =   242
         BeginProperty Column00 
            ColumnWidth     =   2115,213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   599,811
         EndProperty
      EndProperty
   End
   Begin VB.Label lblultimamodifica 
      Caption         =   "25/07/2017"
      Height          =   255
      Left            =   8190
      TabIndex        =   54
      Top             =   6330
      Width           =   1185
   End
   Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL12 
      Height          =   300
      Left            =   5490
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   1710
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   529
      Caption         =   "Ubicazione"
   End
   Begin PRJFW_EDIT.TxtEdit TXT_TIPO_UBICAZIONE 
      Height          =   300
      Left            =   6120
      TabIndex        =   52
      ToolTipText     =   "PZ"
      Top             =   1680
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   529
      Enabled         =   0   'False
      MaxChar         =   2
      Numerico        =   0   'False
      IsDbField       =   0   'False
      Caption         =   "PZ"
      Object.Tag             =   "PZ"
      MaxWidth        =   2
   End
   Begin PRJFW_EDIT.TxtEdit TXT_UBICAZIONE 
      Height          =   300
      Left            =   6480
      TabIndex        =   51
      ToolTipText     =   "PZ"
      Top             =   1680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      Enabled         =   0   'False
      MaxChar         =   12
      Numerico        =   0   'False
      IsDbField       =   0   'False
      Caption         =   "PZ"
      Object.Tag             =   "PZ"
      MaxWidth        =   12
   End
   Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL3 
      Height          =   300
      Left            =   3960
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   1320
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   529
      Caption         =   "Gr.St.2"
   End
   Begin PRJFW_EDIT.TxtEdit TXT_GRST2 
      Height          =   300
      Left            =   4560
      TabIndex        =   49
      ToolTipText     =   "Codice Sottogruppo"
      Top             =   1320
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   529
      Enabled         =   0   'False
      MaxChar         =   4
      Numerico        =   0   'False
      IsDbField       =   0   'False
      Caption         =   "Codice Sottogruppo"
      Object.Tag             =   "Codice Sottogruppo"
      MaxWidth        =   4
      CanRequired     =   0   'False
   End
   Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL10 
      Height          =   300
      Left            =   8820
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   4140
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Caption         =   "Totale Scarichi"
   End
   Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL9 
      Height          =   300
      Left            =   6030
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   4110
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Caption         =   "Totale Carichi"
   End
   Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL8 
      Height          =   300
      Left            =   12030
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Caption         =   "Costo medio"
   End
   Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL7 
      Height          =   300
      Left            =   8820
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   3780
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Caption         =   "Data Ul. Scarico"
   End
   Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL6 
      Height          =   300
      Left            =   6030
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   3780
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Caption         =   "Data Ul. Carico"
   End
   Begin PRJFW_EDIT.TxtEdit TXT_TOTSCARICHI 
      Height          =   300
      Left            =   10020
      TabIndex        =   41
      ToolTipText     =   "PZ"
      Top             =   4110
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      Enabled         =   0   'False
      MaxChar         =   12
      Numerico        =   0   'False
      IsDbField       =   0   'False
      Caption         =   "PZ"
      Object.Tag             =   "PZ"
   End
   Begin PRJFW_EDIT.TxtEdit TXT_TOTCARICHI 
      Height          =   300
      Left            =   7260
      TabIndex        =   40
      ToolTipText     =   "PZ"
      Top             =   4110
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      Enabled         =   0   'False
      MaxChar         =   12
      Numerico        =   0   'False
      IsDbField       =   0   'False
      Caption         =   "PZ"
      Object.Tag             =   "PZ"
   End
   Begin PRJFW_EDIT.TxtEdit TXT_RICMEDIA 
      Height          =   300
      Left            =   12990
      TabIndex        =   39
      ToolTipText     =   "PZ"
      Top             =   3810
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      Enabled         =   0   'False
      MaxChar         =   12
      Object.Visible         =   0   'False
      Numerico        =   0   'False
      IsDbField       =   0   'False
      Caption         =   "PZ"
      Object.Tag             =   "PZ"
   End
   Begin PRJFW_EDIT.TxtEdit TXT_DATAULSCA 
      Height          =   300
      Left            =   10020
      TabIndex        =   38
      ToolTipText     =   "PZ"
      Top             =   3750
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      Enabled         =   0   'False
      MaxChar         =   12
      Numerico        =   0   'False
      IsDbField       =   0   'False
      Caption         =   "PZ"
      Object.Tag             =   "PZ"
   End
   Begin PRJFW_EDIT.TxtEdit TXT_DATAULCA 
      Height          =   300
      Left            =   7260
      TabIndex        =   37
      ToolTipText     =   "PZ"
      Top             =   3750
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      Enabled         =   0   'False
      MaxChar         =   12
      Numerico        =   0   'False
      IsDbField       =   0   'False
      Caption         =   "PZ"
      Object.Tag             =   "PZ"
   End
   Begin PRJFW_GROUPBOX.TMS_GROUPBOX TMS_GROUPBOX4 
      Height          =   1005
      Left            =   5880
      Top             =   3660
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   1773
   End
   Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL5 
      Height          =   300
      Left            =   90
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5610
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   529
      Caption         =   "Costo di acquisto"
   End
   Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL4 
      Height          =   300
      Left            =   60
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   3750
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   529
      Caption         =   "Listini di vendita"
   End
   Begin PRJFW_GROUPBOX.TMS_GROUPBOX TMS_GROUPBOX3 
      Height          =   765
      Left            =   0
      Top             =   5520
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1349
   End
   Begin PRJFW_GROUPBOX.TMS_GROUPBOX TMS_GROUPBOX2 
      Height          =   1005
      Left            =   0
      Top             =   3660
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1773
   End
   Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL2 
      Height          =   300
      Left            =   120
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1740
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   529
      Caption         =   "Art. Fornitore"
   End
   Begin PRJFW_EDIT.TxtEdit TXT_INESAUR 
      Height          =   300
      Left            =   8610
      TabIndex        =   24
      ToolTipText     =   "PZ"
      Top             =   930
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   529
      Enabled         =   0   'False
      MaxChar         =   20
      Numerico        =   0   'False
      IsDbField       =   0   'False
      Caption         =   "PZ"
      Object.Tag             =   "PZ"
      MaxWidth        =   20
      Allineamento    =   2
   End
   Begin PRJFW_EDIT.TxtEdit TXT_DESGRUSTAT1 
      Height          =   300
      Left            =   5520
      TabIndex        =   23
      ToolTipText     =   "PZ"
      Top             =   1320
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   529
      Enabled         =   0   'False
      MaxChar         =   20
      Numerico        =   0   'False
      IsDbField       =   0   'False
      Caption         =   "PZ"
      Object.Tag             =   "PZ"
      MaxWidth        =   20
   End
   Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL1 
      Height          =   300
      Left            =   5280
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   960
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   529
      Caption         =   "PZ"
   End
   Begin PRJFW_EDIT.TxtEdit TXT_PZ 
      Height          =   300
      Left            =   5520
      TabIndex        =   21
      ToolTipText     =   "PZ"
      Top             =   930
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      Enabled         =   0   'False
      MaxChar         =   12
      Numerico        =   0   'False
      IsDbField       =   0   'False
      Caption         =   "PZ"
      Object.Tag             =   "PZ"
      MaxWidth        =   12
   End
   Begin PRJFW_EDIT.TxtEdit TXT_DESCFAM 
      Height          =   300
      Left            =   12240
      TabIndex        =   19
      ToolTipText     =   "Descrizione Famiglia"
      Top             =   1230
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   529
      Enabled         =   0   'False
      MaxChar         =   80
      Object.Visible         =   0   'False
      Numerico        =   0   'False
      Carattere       =   0   'False
      IsDbField       =   0   'False
      Caption         =   "Descrizione Famiglia"
      Object.Tag             =   "Descrizione Famiglia"
      MaxWidth        =   20
   End
   Begin PRJFW_TmsLabel.TMS_LABEL LBL_FAM 
      Height          =   300
      Left            =   150
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      Caption         =   "Fm-Sfm-Gr-Sg"
   End
   Begin PRJFW_COMBOBOX.TMS_COMBO CMB_TIPOART 
      Height          =   315
      Left            =   11910
      TabIndex        =   12
      Top             =   1770
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   556
      Enabled         =   0   'False
      MaxChar         =   20
      Object.Visible         =   0   'False
      IsDbField       =   0   'False
      DbCol           =   0
      Caption         =   "Tipo Articolo"
      Object.Tag             =   "Tipo Articolo"
   End
   Begin PRJFW_EDIT.TxtEdit TXT_TIPOPROD 
      Height          =   300
      Left            =   12270
      TabIndex        =   11
      ToolTipText     =   "Tipo Prodotto"
      Top             =   570
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   529
      Enabled         =   0   'False
      MaxChar         =   40
      Object.Visible         =   0   'False
      Numerico        =   0   'False
      Carattere       =   0   'False
      IsDbField       =   0   'False
      Caption         =   "Tipo Prodotto"
      Object.Tag             =   "Tipo Prodotto"
      MaxWidth        =   16
   End
   Begin PRJFW_EDIT.TxtEdit TXT_SGRUP 
      Height          =   300
      Left            =   3345
      TabIndex        =   10
      ToolTipText     =   "Codice Sottogruppo"
      Top             =   1320
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   529
      Enabled         =   0   'False
      MaxChar         =   4
      Numerico        =   0   'False
      IsDbField       =   0   'False
      Caption         =   "Codice Sottogruppo"
      Object.Tag             =   "Codice Sottogruppo"
      MaxWidth        =   4
      CanRequired     =   0   'False
   End
   Begin PRJFW_EDIT.TxtEdit TXT_GRUP 
      Height          =   300
      Left            =   2730
      TabIndex        =   9
      ToolTipText     =   "Codice Gruppo"
      Top             =   1320
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   529
      Enabled         =   0   'False
      MaxChar         =   4
      Numerico        =   0   'False
      IsDbField       =   0   'False
      Caption         =   "Codice Gruppo"
      Object.Tag             =   "Codice Gruppo"
      MaxWidth        =   4
      CanRequired     =   0   'False
   End
   Begin PRJFW_EDIT.TxtEdit TXT_SFAM 
      Height          =   300
      Left            =   2115
      TabIndex        =   8
      ToolTipText     =   "Codice Sottofamiglia"
      Top             =   1320
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   529
      Enabled         =   0   'False
      MaxChar         =   4
      Numerico        =   0   'False
      IsDbField       =   0   'False
      Caption         =   "Codice Sottofamiglia"
      Object.Tag             =   "Codice Sottofamiglia"
      MaxWidth        =   4
      CanRequired     =   0   'False
   End
   Begin PRJFW_EDIT.TxtEdit TXT_FAM 
      Height          =   300
      Left            =   1500
      TabIndex        =   7
      ToolTipText     =   "Codice Famiglia"
      Top             =   1320
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   529
      Enabled         =   0   'False
      MaxChar         =   4
      Numerico        =   0   'False
      IsDbField       =   0   'False
      Caption         =   "Codice Famiglia"
      Object.Tag             =   "Codice Famiglia"
      MaxWidth        =   4
      CanRequired     =   0   'False
   End
   Begin PRJFW_EDIT.TxtEdit TXT_UM2 
      Height          =   300
      Left            =   12180
      TabIndex        =   6
      ToolTipText     =   "UM 2"
      Top             =   2940
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
      Enabled         =   0   'False
      MaxChar         =   3
      Object.Visible         =   0   'False
      Numerico        =   0   'False
      IsDbField       =   0   'False
      Caption         =   "UM 2"
      Object.Tag             =   "UM 2"
      MaxWidth        =   5
   End
   Begin PRJFW_EDIT.TxtEdit TXT_UM1 
      Height          =   300
      Left            =   3900
      TabIndex        =   5
      ToolTipText     =   "UM 1"
      Top             =   930
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
      Enabled         =   0   'False
      MaxChar         =   3
      Numerico        =   0   'False
      IsDbField       =   0   'False
      Caption         =   "UM 1"
      Object.Tag             =   "UM 1"
      MaxWidth        =   5
   End
   Begin PRJFW_TmsLabel.TMS_LABEL LBL_UM 
      Height          =   300
      Left            =   3480
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   930
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   529
      Caption         =   "UM 1"
   End
   Begin PRJFW_COMBOBOX.TMS_COMBO CMB_TIPOQTA 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   900
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      Enabled         =   0   'False
      MaxChar         =   16
      Obbligatorio    =   -1  'True
      IsDbField       =   0   'False
      DbCol           =   0
      Caption         =   "Tipo quantità"
      Object.Tag             =   "Tipo quantità"
   End
   Begin PRJFW_TmsLabel.TMS_LABEL LBL_QTA 
      Height          =   300
      Left            =   150
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   930
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   529
      Caption         =   "Tipo quantità"
   End
   Begin PRJFW_TmsLabel.TMS_LABEL LBL_CODART 
      Height          =   300
      Left            =   150
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   180
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   529
      Caption         =   "Articolo"
   End
   Begin PRJFW_GROUPBOX.TMS_GROUPBOX TMS_GROUPBOX1 
      Height          =   825
      Left            =   0
      Top             =   30
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   1455
   End
   Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL11 
      Height          =   300
      Left            =   90
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Caption         =   "Listini di acquisto"
   End
   Begin PRJFW_GROUPBOX.TMS_GROUPBOX TMS_GROUPBOX5 
      Height          =   855
      Left            =   0
      Top             =   4650
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1508
   End
   Begin PRJFW_GROUPBOX.TMS_GROUPBOX TMS_GROUPBOX6 
      Height          =   1305
      Left            =   0
      Top             =   810
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   2302
   End
   Begin PRJFW_GROUPBOX.TMS_GROUPBOX TMS_GROUPBOX7 
      Height          =   465
      Left            =   8460
      Top             =   870
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   820
   End
   Begin PRJFW_GROUPBOX.TMS_GROUPBOX TMS_GROUPBOX8 
      Height          =   1155
      Left            =   8460
      Top             =   1350
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   2037
   End
   Begin PRJFW_GROUPBOX.TMS_GROUPBOX TMS_GROUPBOX9 
      Height          =   1185
      Left            =   8460
      Top             =   2490
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   2090
   End
End
Attribute VB_Name = "FRMMG_SITGIACENZE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Gcls_Global                              As CLSFW_Global
Public Gcls_Log                                 As CLSFW_SrvLog
Public Gcon_Connect                             As ADODB.Connection
Public Gcls_Connect                             As New CLSFW_SetConnect
Public Gstr_Connect                             As String

Public ActiveInterface                          As Cinterface
Public ActiveClass                              As CLSMG_INQLIS
Private pbol_alreadyloaded                      As Boolean

Public Gcls_RecSet_SitGiacenze                  As New CLSFW_Recordset
Public WithEvents Grst_SitGiacenze              As ADODB.Recordset
Attribute Grst_SitGiacenze.VB_VarHelpID = -1
Public WithEvents FME_CCS_SKPROD                As CLSFW_VIRTUALFRAME
Attribute FME_CCS_SKPROD.VB_VarHelpID = -1
Public Gstr_SQL_SitGiacenze                     As String

Public Gstr_DittaCorrente                       As String

Public Prst_Progressivi                         As ADODB.Recordset
Attribute Prst_Progressivi.VB_VarHelpID = -1

'Enzo 200703 - Carico listini vendita e acquisto
Public Grst_RecSet_LI11VEN                    As ADODB.Recordset
Public Grst_RecSet_LI11_appendVEN             As ADODB.Recordset
Public Grst_RecSet_LI11ACQ                    As ADODB.Recordset
Public Grst_RecSet_LI11_appendACQ             As ADODB.Recordset
Public Grst_RecSet_LI11ACQ_TOT                As ADODB.Recordset
Public Grst_RecSet_LI11_appendACQ_TOT         As ADODB.Recordset


'Enzo 200703 - Carico data ultimo carico/scarico
Public Prst_DataCar                        As ADODB.Recordset

'Enzo 200703 - Calcolo prezzi
Public Gcls_CalcoloPrezzi                  As MGBO_PREZZI.CLSMG_CALCPRNETTO

'Enzo 200703 - Pezzi per confezione preferenziale
Public RecDatiAppoggio                             As ADODB.Recordset



Private Pcls_GridFormat                         As CLSFW_DataGrid

'Private WithEvents FrmDocumenti                 As FRMMG_VISDOC
'Private WithEvents FrmODL                       As FRMMG_VISODL
'Private WithEvents FrmDispo                     As FRMMG_DISPO
'Private WithEvents FrmImprod                    As FRMMG_IMPPROD
'Private WithEvents FrmScortaProd                As FRMMG_SCORTAPROD
'Private WithEvents FrmPreImpCli                 As FRMMG_PREIMPCLI

Private Pcls_Decode                             As MGBO_LOOKUPDECODE.CLSMG_DECODE
Private Pcls_Lookup                             As MGBO_LOOKUPDECODE.CLSMG_LOOKUP
Private Old_Articolo                            As String
Private OnUnload                                As Boolean
Private IsLOaded                                As Boolean
Private OnClicLookUp                            As Boolean
Private ClickNuovo                              As Boolean
Private ValidateArticolo                        As Boolean
Private ValidateOpzione                         As Boolean
Private PbolLookupArticForn                     As Boolean
Private PbolLookupArticCli                      As Boolean
Private WithEvents Pcls_Partitario              As CLSMG_INTPART
Attribute Pcls_Partitario.VB_VarHelpID = -1

#If Not GAMMA_SPRINT Then
    Private WithEvents Pcls_DispoProd           As CLSPD_CCS_ESPLGIA
Attribute Pcls_DispoProd.VB_VarHelpID = -1
    Private WithEvents Pcls_CicloLavorazione    As CLSPD_GESCICLI
Attribute Pcls_CicloLavorazione.VB_VarHelpID = -1
    Private Pcls_Connect_Produzione             As PDBO_LOOKUPDECODE.CLSPD_CONNECT
#End If

Private WithEvents Cls_ConnectMagazzino         As MGBO_LOOKUPDECODE.CLSMG_CONNECT
Attribute Cls_ConnectMagazzino.VB_VarHelpID = -1
Private WithEvents Pcls_ArtClienti              As CLSMG_ARTCLI
Attribute Pcls_ArtClienti.VB_VarHelpID = -1
Private WithEvents Pcls_ArtFornitori            As CLSMG_ARTFOR
Attribute Pcls_ArtFornitori.VB_VarHelpID = -1

Private PermAnagrArt                            As Variant
Private PermPartitario                          As Variant
Private PermCicloLavor                          As Variant
Private PermDisponibilità                       As Variant
Private PermArtClienti                          As Variant
Private PermArtFornitori                        As Variant
Private NumProg                                 As Integer
Private cls_datagrid                            As CLSFW_DataGrid
Private WithEvents Pstd_Format                  As StdDataFormat
Attribute Pstd_Format.VB_VarHelpID = -1
Private WithEvents Pstd_FormatDEP               As StdDataFormat
Attribute Pstd_FormatDEP.VB_VarHelpID = -1
Private Pint_LookupPers                         As Integer
Private WithEvents Pcls_SkPrezzi                As CLSMG_SCHEDAPRZART
Attribute Pcls_SkPrezzi.VB_VarHelpID = -1
Private PermSkPrezzi                            As Variant
Private pvarDecimali                            As Variant

'Public Function CcsPermessiPrezzi_MENU(DLL_Classe As String) As Variant
'On Error GoTo Err
'Dim stringa                 As String
'Dim rstPermessi             As ADODB.Recordset
'Dim Utente                  As String
'Dim Gruppo                  As String
'Dim FlgGruppo               As String
'
'    Utente = ActiveInterface.ClsGlobal.Gcls_UtenteCorrente.Codice
'    Gruppo = ActiveInterface.ClsGlobal.Gcls_UtenteCorrente.Gruppo
'    FlgGruppo = ActiveInterface.ClsGlobal.Gcls_UtenteCorrente.FlagGrp
'
'    If Not rstPermessi Is Nothing Then
'        If rstPermessi.State = adStateOpen Then
'            rstPermessi.Close
'        End If
'    Else
'        Set rstPermessi = New ADODB.Recordset
'    End If
'
'    stringa = "SELECT * FROM FW50_MODMENU  WITH (NOLOCK) "
'    stringa = stringa & " INNER JOIN FW52_RVKMENU  WITH (NOLOCK) ON FW52_IDVOCETS_FW50 = FW50_IDVOCETS "
'    stringa = stringa & " WHERE FW50_NOME = '" & DLL_Classe & "' "
'
'    If FlgGruppo = "1" Then
'        stringa = stringa & " AND FW52_GRUPPO_FW06 = '" & Gruppo & "' AND (FW52_UTENTE_FW07 = '' OR FW52_UTENTE_FW07 IS NULL) "
'    Else
'        stringa = stringa & " AND FW52_GRUPPO_FW06 = '" & Gruppo & "' AND FW52_UTENTE_FW07 = '" & Utente & "' "
'    End If
'
'    stringa = stringa & " AND FW52_FLGABIL = 0"
'
'    Set rstPermessi = Gcon_Connect.Execute(stringa)
'
'    If rstPermessi.EOF Then
'        CcsPermessiPrezzi_MENU = True
'    Else
'        CcsPermessiPrezzi_MENU = False
'    End If
'
'    rstPermessi.Close
'
'    Exit Function
'Err:
'  Set Gcls_Log.vbError = Err
'  Set Gcls_Log.ADOError = Gcon_Connect.Errors
'  If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_VISDOC.CcsPermessiPrezzi_MENU") = 1 Then
'     Unload Me
'  Else
'     Resume Next
'  End If
'End Function
'
'Private Sub PreImpegnoCliente()
'    On Error GoTo Err
'
'    Set FrmPreImpCli = New FRMMG_PREIMPCLI
'    Set FrmPreImpCli.Gcon_Connect = Gcon_Connect
'    Set FrmPreImpCli.Gcls_Log = Gcls_Log
'    Set FrmPreImpCli.ActiveInterface = ActiveInterface
'    FrmPreImpCli.Gstr_Connect = Gstr_Connect
'    FrmPreImpCli.Gstr_DittaCorrente = Gstr_DittaCorrente
'    FrmPreImpCli.Articolo = RTrimN(TXT_CODART.Text)
'    FrmPreImpCli.Variante = RTrimN(TXT_OPZIONE.Text)
'    FrmPreImpCli.Descrizione = RTrimN(TXT_DESCART.Text)
'    If ActiveInterface.WindowModal Then
'        FrmPreImpCli.Show vbModal
'    Else
'        Me.Hide
'        FrmPreImpCli.Show vbModeless
'    End If
'
'    Exit Sub
'
'Err:
'  Set Gcls_Log.vbError = Err
'  Set Gcls_Log.ADOError = Gcon_Connect.Errors
'  If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.PreImpegnoCliente") = 1 Then
'     Unload Me
'  Else
'     Resume Next
'  End If
'End Sub

Private Sub CMD_ELABORA_ButtonClick()

    On Error GoTo Err
    
    Decimali
    If RTrimN(TXT_CODART.Text) <> "" And TXT_CODART.IsValid Then
        TXT_CODART.Enabled = False
        TXT_OPZIONE.Enabled = False
        CMB_TIPOQTA.Enabled = False
        CMD_ELABORA.Enabled = False
        Call RiempioDati(RTrimN(Grst_SitGiacenze.Fields("MG66_CODART").Value), "")
        Call Psub_Elabora(RTrimN(TXT_CODART.Text), RTrimN(TXT_OPZIONE.Text))
    Else
        MsgBox "Campo obbligatorio mancante!", vbCritical, "Informazione"
        TXT_CODART.SetTextFocus
    End If
    
    Exit Sub

Err:
  Set Gcls_Log.vbError = Err
  Set Gcls_Log.ADOError = Gcon_Connect.Errors
  If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.CMD_ELABORA_ButtonClick") = 1 Then
     Unload Me
  Else
     Resume Next
  End If

End Sub

'Private Sub CMD_PREIMPCLI_ButtonClick()
'    On Error Resume Next
'    Call PreImpegnoCliente
'    Err.Clear
'End Sub

'Private Sub Disponibile()
'    On Error GoTo Err
'
'    'Richiamo la FORM x visualizzare i documenti
'    Set FrmDispo = New FRMMG_DISPO
'    Set FrmDispo.Gcon_Connect = Gcon_Connect
'    Set FrmDispo.Gcls_Log = Gcls_Log
'    Set FrmDispo.ActiveInterface = ActiveInterface
'    FrmDispo.Gstr_Connect = Gstr_Connect
'    FrmDispo.Gstr_DittaCorrente = Gstr_DittaCorrente
'    FrmDispo.Articolo = RTrimN(TXT_CODART.Text)
'    FrmDispo.Variante = RTrimN(TXT_OPZIONE.Text)
'    FrmDispo.Descrizione = RTrimN(TXT_DESCART.Text)
'    If ActiveInterface.WindowModal Then
'        FrmDispo.Show vbModal
'    Else
'        Me.Hide
'        FrmDispo.Show vbModeless
'    End If
'
'    Exit Sub
'
'Err:
'  Set Gcls_Log.vbError = Err
'  Set Gcls_Log.ADOError = Gcon_Connect.Errors
'  If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.Disponibile") = 1 Then
'     Unload Me
'  Else
'     Resume Next
'  End If
'End Sub

'Private Sub CHK_MOV_Click()
'    On Error GoTo Err
'
'    Call Psub_Elabora(RTrimN(TXT_CODART.Text), RTrimN(TXT_OPZIONE.Text))
'
'    Exit Sub
'
'Err:
'  Set Gcls_Log.vbError = Err
'  Set Gcls_Log.ADOError = Gcon_Connect.Errors
'  If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.CHK_MOV_Click") = 1 Then
'     Unload Me
'  Else
'     Resume Next
'  End If
'End Sub

'Private Sub CMD_DISPO_ButtonClick()
'    On Error Resume Next
'    Call Disponibile
'    Err.Clear
'End Sub

'Private Sub RicercaOCOF(TipoDoc As Integer)
'    On Error GoTo Err:
'
'    'Richiamo la FORM x visualizzare i documenti
'    Set FrmDocumenti = New FRMMG_VISDOC
'    Set FrmDocumenti.Gcon_Connect = Gcon_Connect
'    Set FrmDocumenti.Gcls_Log = Gcls_Log
'    Set FrmDocumenti.ActiveInterface = ActiveInterface
'    FrmDocumenti.Gstr_Connect = Gstr_Connect
'    FrmDocumenti.Gstr_DittaCorrente = Gstr_DittaCorrente
'    FrmDocumenti.Articolo = RTrimN(TXT_CODART.Text)
'    FrmDocumenti.Variante = RTrimN(TXT_OPZIONE.Text)
'    FrmDocumenti.Descrizione = RTrimN(TXT_DESCART.Text)
'    If TipoDoc = 21 Then
'        FrmDocumenti.VenAcq = 0
'    Else
'        FrmDocumenti.VenAcq = 1
'    End If
'    FrmDocumenti.TipoDocumento = TipoDoc
'    If ActiveInterface.WindowModal Then
'        FrmDocumenti.Show vbModal
'    Else
'        Me.Hide
'        FrmDocumenti.Show vbModeless
'    End If
'
'    Exit Sub
'
'Err:
'  Set Gcls_Log.vbError = Err
'  Set Gcls_Log.ADOError = Gcon_Connect.Errors
'  If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.RicercaOCOF") = 1 Then
'     Unload Me
'  Else
'     Resume Next
'  End If
'End Sub

'Private Sub CMD_IMPCLI_ButtonClick()
'    Call RicercaOCOF(21)
'End Sub
'
'Private Sub CMD_ORDFOR_ButtonClick()
'    Call RicercaOCOF(22)
'End Sub
'
'Private Sub RicercaODL()
'    On Error GoTo Err
'
'    'Richiamo la FORM x visualizzare i documenti
'    Set FrmODL = New FRMMG_VISODL
'    Set FrmODL.Gcon_Connect = Gcon_Connect
'    Set FrmODL.Gcls_Log = Gcls_Log
'    Set FrmODL.ActiveInterface = ActiveInterface
'    FrmODL.Gstr_Connect = Gstr_Connect
'    FrmODL.Gstr_DittaCorrente = Gstr_DittaCorrente
'    FrmODL.Articolo = RTrimN(TXT_CODART.Text)
'    FrmODL.Variante = RTrimN(TXT_OPZIONE.Text)
'    FrmODL.Descrizione = RTrimN(TXT_DESCART.Text)
'    If ActiveInterface.WindowModal Then
'        FrmODL.Show vbModal
'    Else
'        Me.Hide
'        FrmODL.Show vbModeless
'    End If
'
'    Exit Sub
'
'Err:
'  Set Gcls_Log.vbError = Err
'  Set Gcls_Log.ADOError = Gcon_Connect.Errors
'  If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.RicercaODL") = 1 Then
'     Unload Me
'  Else
'     Resume Next
'  End If
'End Sub
'
'
'Private Sub CMD_ORDPRO_ButtonClick()
'    Call RicercaODL
'End Sub

'Private Sub RicercaImprod()
'    On Error GoTo Err:
'
'    'Richiamo la FORM x visualizzare i documenti
'    Set FrmImprod = New FRMMG_IMPPROD
'    Set FrmImprod.Gcon_Connect = Gcon_Connect
'    Set FrmImprod.Gcls_Log = Gcls_Log
'    Set FrmImprod.ActiveInterface = ActiveInterface
'    FrmImprod.Gstr_Connect = Gstr_Connect
'    FrmImprod.Gstr_DittaCorrente = Gstr_DittaCorrente
'    FrmImprod.Articolo = RTrimN(TXT_CODART.Text)
'    FrmImprod.Variante = RTrimN(TXT_OPZIONE.Text)
'    FrmImprod.Descrizione = RTrimN(TXT_DESCART.Text)
'    If ActiveInterface.WindowModal Then
'        FrmImprod.Show vbModal
'    Else
'        Me.Hide
'        FrmImprod.Show vbModeless
'    End If
'
'    Exit Sub
'
'Err:
'  Set Gcls_Log.vbError = Err
'  Set Gcls_Log.ADOError = Gcon_Connect.Errors
'  If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.RicercaImprod") = 1 Then
'     Unload Me
'  Else
'     Resume Next
'  End If
'End Sub
'
'Private Sub CMD_IMPPROD_ButtonClick()
'    Call RicercaImprod
'End Sub

Private Sub CMD_NUOVO_ButtonClick()
    On Error GoTo Err:
    
    'Metto il VirtualFrame in stato inserimento x attivare l'inserimento
    'di un nuovo codice da visualizzare
    If TXT_CODART.IsValid Then
        'FME_CCS_SKPROD.Status = tsInsert
    Else
        TXT_CODART.Text = ""
        TXT_OPZIONE.Text = ""
        TXT_CODART.SetTextFocus
    End If
    'Disattivo il messaggio a richiesta di aggiornare i dati modificati
    FME_CCS_SKPROD.MsgOnUpdate = False
    
    FME_CCS_SKPROD.Status = tsInsert
    
    Call Psub_Reinizializza

'    CMD_DISPO.Enabled = False
'    CMD_IMPCLI.Enabled = False
'    CMD_IMPPROD.Enabled = False
'    CMD_ORDFOR.Enabled = False
'    CMD_ORDPRO.Enabled = False
'    CMD_PREIMPCLI.Enabled = False
'    CMD_COLLEGAMENTI.Enabled = False
        
    TXT_CODART.Text = ""
    TXT_CODART.Enabled = True
    TXT_CODART.SetFocus
    TXT_OPZIONE.Enabled = True
    CMB_TIPOQTA.Enabled = True
    
    Call ReinizializzaVirtualFrame
    
    'Enzo 200703 Pulisci campi nuovi
    TXT_PZ.Text = ""
    TXT_INESAUR.Text = ""
    TXT_DESGRUSTAT1.Text = ""
    TXT_CODARTFOR.Text = ""
    TXT_CODARTSOST.Text = ""
    TXT_CODARTSOSTDES.Text = ""
    TXT_DESCARTCF.Text = ""
    TXT_RICMEDIA.Text = ""
    
    
    'Variabile x sapere il passaggio dal tasto NUOVO
    ClickNuovo = True
    
    'Verifico se devo attivare la variante
'    If ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsParProd.ModalitaGestioneVarianti = tsConfiguratore _
'    Or ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsParProd.ModalitaGestioneVarianti = tsVariantiArichiesta Then
'        TXT_OPZIONE.Enabled = True
'    Else
'        TXT_OPZIONE.Enabled = False
'    End If
    CMD_ELABORA.Enabled = True
    
    Exit Sub

Err:
  Set Gcls_Log.vbError = Err
  Set Gcls_Log.ADOError = Gcon_Connect.Errors
  If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.ButtonClick") = 1 Then
     Unload Me
  Else
     Resume Next
  End If
End Sub

'Private Sub CMD_COLLEGAMENTI_MenuItemClick(ByVal Index As Integer, ByVal Key As String, ByVal Caption As String)
'    On Error GoTo Err
'
'    Select Case Key
'        Case "Key_Anagrafica"
'            Call InvocaAnagArticoli
'        Case "Key_Partitario"
'            Call InvocaPartitario
'        Case "Key_Disponibilità"
'            Call InvocaDispoProd
'        Case "Key_DatiScorteProd"
'            Call VisDatiScortaProduzione
'        Case "Key_CicloLavorazione"
'            Call InvocaCicloLavorazione
'        Case "Key_ArtClienti"
'            Call InvocaArtClienti
'        Case "Key_ArtFornitori"
'            Call InvocaArtFornitori
'        Case "Key_SkPrezziAcq"
'            Call InvocaSkPrezzi(1)
'        Case "Key_SkPrezziVen"
'            Call InvocaSkPrezzi(0)
'#If Not GAMMA_SPRINT Then
'        Case "Key_GiacCLav"
'            Call InvocaVerificaGiacenzeCLavoro
'#End If
'    End Select
'
'   Exit Sub
'
'Err:
'  Set Gcls_Log.vbError = Err
'  Set Gcls_Log.ADOError = Gcon_Connect.Errors
'  If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.CMD_COLLEGAMENTI_MenuItemClick") = 1 Then
'     Unload Me
'  Else
'     Resume Next
'  End If
'End Sub

Private Sub RiempioDati(CodArt As String, Variante As String)
    Dim sql                                 As String
    Dim RecDati                             As ADODB.Recordset
    Dim Str_Descrizione                     As String

    'Enzo 200703 - Pezzi per confezione preferenziale
    Dim sqlAppoggio                                 As String

    On Error GoTo Err

'   Set RecDati = New ADODB.Recordset

    'Query x leggere i dati con cui rimpire le informazioni iniziali
    sql = "SELECT MG66_UM1,MG66_UM2,MG66_FAM_MG53,MG66_SFAM_MG54," _
        & "MG66_GRUPPO_MG55,MG66_SGRUPPO_MG56,MG66_GRUSTAT2_MG75, PD18_INDTIPOART," _
        & "PD20_DESCR,MG53_DESCRFAM,MG54_DESCRSFAM,MG55_DESCRGRUPPO," _
        & "MG56_DESCRSGRUPPO,MG87_DESCART "
    
    'Enzo 200703 - Descrizione estesa articolo
    sql = sql & ", MG87_DESCARTEST " 'MG66_INDSTATO_MG6W
    
    
    'Enzo 200703 - Descrizione gruppo statistico 1
    sql = sql & ", MG66_GRUSTAT1_MG74, MG74_DESGRUSTAT1"
    
    'Enzo 200703 - Descrizione articolo in esaurimento
    sql = sql & ", MG66_INDSTATO_MG6W, MG6W_DESCR, MG7A_UBICAZFIX, MG97_COORDIN1" 'MG66_INDSTATO_MG6W
    
    
    sql = sql & " FROM MG66_ANAGRART  WITH (NOLOCK) " _
        & "LEFT OUTER JOIN PD18_ARTPROD  WITH (NOLOCK) " _
        & "    ON MG66_DITTA_CG18 = PD18_DITTA_CG18 " _
        & "    AND MG66_CODART = PD18_CODART_MG66 " _
        & "LEFT OUTER JOIN PD20_TIPOPROD  WITH (NOLOCK) " _
        & "    ON PD18_DITTA_CG18 = PD20_DITTA_CG18 " _
        & "    AND PD18_TIPOPROD_PD20 = PD20_CODPROD " _
        & "LEFT OUTER JOIN MG53_FAMIGLIE  WITH (NOLOCK) " _
        & "    ON MG66_DITTA_CG18 = MG53_DITTA_CG18 " _
        & "    AND MG66_FAM_MG53 = MG53_CODFAM " _
        & "LEFT OUTER JOIN MG54_SOTTOFAM  WITH (NOLOCK) " _
        & "    ON MG66_DITTA_CG18 = MG54_DITTA_CG18 " _
        & "    AND MG66_FAM_MG53 = MG54_CODFAM_MG53 " _
        & "    AND MG66_SFAM_MG54 = MG54_CODSFAM " _
        & "LEFT OUTER JOIN MG55_GRUPPI  WITH (NOLOCK) " _
        & "    ON MG66_DITTA_CG18   = MG55_DITTA_CG18 " _
        & "    AND MG66_FAM_MG53    = MG55_CODFAM_MG53 " _
        & "    AND MG66_SFAM_MG54   = MG55_CODSFAM_MG54 " _
        & "    AND MG66_GRUPPO_MG55 = MG55_CODGRUPPO "

     'Enzo 200703 - Descrizione gruppo statistico 1
     sql = sql & " LEFT OUTER JOIN MG74_GRUSTAT1  " _
        & "    ON MG66_DITTA_CG18     = MG74_DITTA_CG18 " _
        & "    AND MG66_GRUSTAT1_MG74 = MG74_CODGRUSTAT1 "
     
     sql = sql & "LEFT OUTER JOIN MG56_SOTTOGRUPPI  WITH (NOLOCK) " _
     & "    ON MG66_DITTA_CG18 = MG56_DITTA_CG18 " _
     & "    AND MG66_FAM_MG53 = MG56_CODFAM_MG53 " _
     & "    AND MG66_SFAM_MG54 = MG56_CODSFAM_MG54 " _
     & "    AND MG66_GRUPPO_MG55 = MG56_CODGRUPPO_MG55 " _
     & "    AND MG66_SGRUPPO_MG56 = MG56_CODSGRUPPO " _
     & "LEFT OUTER JOIN MG87_ARTDESC  WITH (NOLOCK) " _
     & "    ON MG66_DITTA_CG18 = MG87_DITTA_CG18 " _
     & "    AND MG66_CODART = MG87_CODART_MG66 " _
     & "    AND MG87_OPZIONE_MG5E = '" & Variante & "' " _
     & "    AND MG87_LINGUA_MG52 = '' "
     
     
     sql = sql & " LEFT OUTER JOIN MG7A_UBICAZARTFIX " _
            & " ON MG66_ANAGRART.MG66_DITTA_CG18 = MG7A_UBICAZARTFIX.MG7A_DITTA_CG18 " _
            & "  AND MG66_ANAGRART.MG66_CODART = MG7A_UBICAZARTFIX.MG7A_CODART_MG66 " _
            & " LEFT OUTER JOIN MG97_UBICAZ " _
            & "  ON MG7A_UBICAZARTFIX.MG7A_UBICAZFIX = MG97_UBICAZ.MG97_UBICAZ " _
            & "  AND MG7A_UBICAZARTFIX.MG7A_CODDEP_MG58 = MG97_UBICAZ.MG97_CODDEP_MG58 " _
            & "  AND MG7A_UBICAZARTFIX.MG7A_DITTA_CG18 = MG97_UBICAZ.MG97_DITTA_CG18 " _

     'Enzo 200907 - Descrizione stato articolo
     sql = sql & " LEFT OUTER JOIN MG6W_STATIART  " _
               & "   ON MG66_INDSTATO_MG6W     = MG6W_INDSTATO "
     
     
     sql = sql & "WHERE MG66_DITTA_CG18 = " & Gstr_DittaCorrente _
     & " AND MG66_CODART = '" & CodArt & "'"

    Set RecDati = Gcon_Connect.Execute(sql, , adCmdText)

    If RecDati.EOF = False Then
        TXT_UM1.Text = RecDati.Fields("MG66_UM1").Value
        TXT_UM2.Text = RecDati.Fields("MG66_UM2").Value
        TXT_FAM.Text = RecDati.Fields("MG66_FAM_MG53").Value
        TXT_SFAM.Text = RecDati.Fields("MG66_SFAM_MG54").Value
        TXT_GRUP.Text = RecDati.Fields("MG66_GRUPPO_MG55").Value
        TXT_SGRUP.Text = RecDati.Fields("MG66_SGRUPPO_MG56").Value
        TXT_GRST2.Text = NVL(RecDati.Fields("MG66_GRUSTAT2_MG75").Value)
        TXT_UBICAZIONE.Text = NVL(RecDati.Fields("MG7A_UBICAZFIX").Value)
        Select Case NVL(RecDati.Fields("MG97_COORDIN1").Value)
        Case 1
        TXT_TIPO_UBICAZIONE.Text = "MU"
        Case 0
        TXT_TIPO_UBICAZIONE.Text = "MA"
        Case Else
        TXT_TIPO_UBICAZIONE.Text = "ZZ"
        End Select
        
        TXT_TIPOPROD.Text = RecDati.Fields("PD20_DESCR").Value
        If Not IsNull(RecDati.Fields("PD18_INDTIPOART").Value) Then
            CMB_TIPOART.Text = RecDati.Fields("PD18_INDTIPOART").Value
        Else
            CMB_TIPOART.Text = 4
        End If
        TXT_DESCART.Text = RecDati.Fields("MG87_DESCART").Value
        
        'Enzo 200703 - Descrizione estesa articolo
        TXT_DESCARTEST.Text = RecDati.Fields("MG87_DESCARTEST").Value
        
        'Enzo 200703 - Descrizione gruppo statistico 1
        TXT_DESGRUSTAT1.Text = RecDati.Fields("MG74_DESGRUSTAT1").Value
        
        'Enzo 200907 - Descrizione stato articolo da tabella
        ' INIZIO *******************************************************
        'Enzo 200703 - Descrizione articolo in esaurimento
'        Select Case RecDati.Fields("MG66_INDSTATO_MG6W").Value
'        Case 50
'          TXT_INESAUR.Text = ""
'        Case 60
'          TXT_INESAUR.Text = "In Esaurimento "
'        Case 90
'          TXT_INESAUR.Text = "Dismesso"
'        Case Else
'          TXT_INESAUR.Text = ""
'        End Select
        
'        If RecDati.Fields("MG66_INDSTATO_MG6W").Value = 1 Then  'MG66_INDSTATO_MG6W
'          TXT_INESAUR.Text = "  *** ESAURITO ***"
'        Else
'          TXT_INESAUR.Text = ""
'        End If
        TXT_INESAUR.Text = RecDati.Fields("MG6W_DESCR").Value
        
        
        ' FINE *******************************************************
        
        'Enzo 200703 - Pezzi per confezione preferenziale
        ' INIZIO *******************************************************
        sqlAppoggio = "SELECT TOP 1 * FROM MG68_CONFART"
        sqlAppoggio = sqlAppoggio & " WHERE MG68_DITTA_CG18 = " & Gstr_DittaCorrente
        sqlAppoggio = sqlAppoggio & " AND MG68_CODART_MG66 = '" & CodArt & "'"
        'sqlAppoggio = sqlAppoggio & " AND MG68_OPZIONE_MG5E "
        sqlAppoggio = sqlAppoggio & " AND MG68_FLGCONFPREF = 1 "
        
        Set RecDatiAppoggio = Gcon_Connect.Execute(sqlAppoggio, , adCmdText)
        If RecDatiAppoggio.EOF = False Then
          TXT_PZ.Text = RecDatiAppoggio.Fields("MG68_PZCONF").Value
        End If
        
        If Not RecDatiAppoggio Is Nothing Then
            Set RecDatiAppoggio = Nothing
        End If
        ' FINE *******************************************************
        
        'Enzo 200703 - Articolo fornitore
        ' INIZIO *******************************************************
        sqlAppoggio = "SELECT * FROM MG73_ARTCLIFOR"
        sqlAppoggio = sqlAppoggio & " WHERE MG73_DITTA_CG18 = " & Gstr_DittaCorrente
        sqlAppoggio = sqlAppoggio & " AND MG73_CODART_MG66 = '" & CodArt & "'"
        'sqlAppoggio = sqlAppoggio & " AND MG68_OPZIONE_MG5E "
        sqlAppoggio = sqlAppoggio & " ORDER BY MG73_FLGFORPREF DESC "
        
        Set RecDatiAppoggio = Gcon_Connect.Execute(sqlAppoggio, , adCmdText)
        If RecDatiAppoggio.EOF = False Then
          TXT_CODARTFOR.Text = RecDatiAppoggio.Fields("MG73_ARTCLIFOR").Value
          TXT_DESCARTCF.Text = RecDatiAppoggio.Fields("MG73_DESCARTCF").Value
        End If
        
        If Not RecDatiAppoggio Is Nothing Then
            Set RecDatiAppoggio = Nothing
        End If
        ' FINE *******************************************************
        
        
        'Enzo 200703 - Articolo sostitutivo
        ' INIZIO *******************************************************
        sqlAppoggio = "SELECT * "
        sqlAppoggio = sqlAppoggio & " FROM         MG85_ARTSOST LEFT OUTER JOIN"
        sqlAppoggio = sqlAppoggio & "                       MG87_ARTDESC ON MG85_ARTSOST.MG85_DITTA_CG18 = MG87_ARTDESC.MG87_DITTA_CG18 AND"
        sqlAppoggio = sqlAppoggio & "                       MG85_ARTSOST.MG85_CODARTEFF_MG66 = MG87_ARTDESC.MG87_CODART_MG66"
        sqlAppoggio = sqlAppoggio & " WHERE MG85_DITTA_CG18 = " & Gstr_DittaCorrente
        sqlAppoggio = sqlAppoggio & " AND MG85_CODARTEFF_MG66 = '" & CodArt & "'"
        'sqlAppoggio = sqlAppoggio & " AND MG85_OPZIONEEFF_MG5E "
        sqlAppoggio = sqlAppoggio & " AND MG85_DATASOST <= CONVERT(DateTime,'" & Format(Now, "mm/dd/yyyy") & "',101) "
        sqlAppoggio = sqlAppoggio & " ORDER BY MG85_DATASOST DESC, MG85_CODARTSOST_MG66 "
        
        Set RecDatiAppoggio = Gcon_Connect.Execute(sqlAppoggio, , adCmdText)
'        If RecDatiAppoggio.EOF = False Then
'          TXT_CODARTSOST.Text = RecDatiAppoggio.Fields("MG85_CODARTSOST_MG66").Value
'          TXT_CODARTSOSTDES.Text = RecDatiAppoggio.Fields("MG87_DESCART").Value
'        End If
        
        Set GRID_ARTSOST.DataSource = RecDatiAppoggio
        
        If Not RecDatiAppoggio Is Nothing Then
            Set RecDatiAppoggio = Nothing
        End If
        ' FINE *******************************************************
        
        'Enzo 200703 - Articolo alternativo - Se non ha trovato l'articolo sostitutivo
        sqlAppoggio = "SELECT * "
        sqlAppoggio = sqlAppoggio & " FROM  MG84_ARTALTER LEFT OUTER JOIN"
        sqlAppoggio = sqlAppoggio & "       MG87_ARTDESC ON MG84_ARTALTER.MG84_DITTA_CG18 = MG87_ARTDESC.MG87_DITTA_CG18 AND"
        sqlAppoggio = sqlAppoggio & "       MG84_ARTALTER.MG84_CODARTEFF_MG66 = MG87_ARTDESC.MG87_CODART_MG66"
        sqlAppoggio = sqlAppoggio & " WHERE MG84_DITTA_CG18 = " & Gstr_DittaCorrente
        sqlAppoggio = sqlAppoggio & " AND MG84_CODARTEFF_MG66 = '" & CodArt & "'"
        'sqlAppoggio = sqlAppoggio & " AND MG84_OPZIONEEFF_MG5E "
        sqlAppoggio = sqlAppoggio & " ORDER BY MG84_CODRAGALT, MG84_CODARTALT_MG66 "
          
        Set RecDatiAppoggio = Gcon_Connect.Execute(sqlAppoggio, , adCmdText)
'        If RecDatiAppoggio.EOF = False Then
'          TXT_CODARTSOST.Text = RecDatiAppoggio.Fields("MG84_CODARTALT_MG66").Value
'          TXT_CODARTSOSTDES.Text = RecDatiAppoggio.Fields("MG87_DESCART").Value
'        Else
'          TXT_CODARTSOST.Text = ""
'          TXT_CODARTSOSTDES.Text = ""
'        End If

        Set GRID_ARTALT.DataSource = RecDatiAppoggio
        
        If Not RecDatiAppoggio Is Nothing Then
            Set RecDatiAppoggio = Nothing
        End If
 

        If Not RecDatiAppoggio Is Nothing Then
            Set RecDatiAppoggio = Nothing
        End If
        
        
        Str_Descrizione = ""
        If RTrimN(RecDati.Fields("MG53_DESCRFAM").Value) <> "" Then
            Str_Descrizione = RTrimN(RecDati.Fields("MG53_DESCRFAM").Value)
        End If
        If RTrimN(RecDati.Fields("MG54_DESCRSFAM").Value) <> "" Then
            If Str_Descrizione <> "" Then
                Str_Descrizione = Str_Descrizione & "/" & RTrimN(RecDati.Fields("MG54_DESCRSFAM").Value)
            Else
                Str_Descrizione = Str_Descrizione & RTrimN(RecDati.Fields("MG54_DESCRSFAM").Value)
            End If
        End If
        If RTrimN(RecDati.Fields("MG55_DESCRGRUPPO").Value) <> "" Then
            If Str_Descrizione <> "" Then
                Str_Descrizione = Str_Descrizione & "/" & RTrimN(RecDati.Fields("MG55_DESCRGRUPPO").Value)
            Else
                Str_Descrizione = Str_Descrizione & RTrimN(RecDati.Fields("MG55_DESCRGRUPPO").Value)
            End If
        End If
        If RTrimN(RecDati.Fields("MG56_DESCRSGRUPPO").Value) <> "" Then
            If Str_Descrizione <> "" Then
                Str_Descrizione = Str_Descrizione & "/" & RTrimN(RecDati.Fields("MG56_DESCRSGRUPPO").Value)
            Else
                Str_Descrizione = Str_Descrizione & RTrimN(RecDati.Fields("MG56_DESCRSGRUPPO").Value)
            End If
        End If
        TXT_DESCFAM.Text = Str_Descrizione

'        If RTrimN(RecDati.Fields("MG66_SGRUPPO_MG56").Value) > "" Then
'            TXT_DESCFAM.Text = RecDati.Fields("MG56_DESCRSGRUPPO").Value
'        Else
'            If RTrimN(RecDati.Fields("MG66_GRUPPO_MG55").Value) > "" Then
'                TXT_DESCFAM.Text = RecDati.Fields("MG55_DESCRGRUPPO").Value
'            Else
'                If RTrimN(RecDati.Fields("MG66_SFAM_MG54").Value) > "" Then
'                    TXT_DESCFAM.Text = RecDati.Fields("MG54_DESCRSFAM").Value
'                Else
'                    TXT_DESCFAM.Text = RecDati.Fields("MG53_DESCRFAM").Value
'                End If
'            End If
'        End If
    End If

    If Not RecDati Is Nothing Then
        Set RecDati = Nothing
    End If
    
    Exit Sub

Err:
    If Not RecDati Is Nothing Then
        Set RecDati = Nothing
    End If
    Set Gcls_Log.vbError = Err
    Set Gcls_Log.ADOError = Gcon_Connect.Errors
    If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.RiempioDati") = 1 Then
       Unload Me
    Else
       Resume Next
    End If
End Sub

'Private Sub CaricaPgmCollegati()
'    On Error GoTo Err
'
'    CMD_COLLEGAMENTI.AddMenuItem "Partitario", "Key_Partitario"
'
'    #If Not GAMMA_SPRINT Then
'        CMD_COLLEGAMENTI.AddMenuItem "Verifica giacenze conto lavoro", "Key_GiacCLav"
'        CMD_COLLEGAMENTI.AddMenuItem "Ciclo di lavorazione", "Key_CicloLavorazione"
'    #End If
'
'    CMD_COLLEGAMENTI.AddMenuItem "Anagrafica articoli", "Key_Anagrafica"
'
'    #If Not GAMMA_SPRINT Then
'        CMD_COLLEGAMENTI.AddMenuItem "Disponibilità esploso", "Key_Disponibilità"
'        If ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.FlgProduzione <> 1 Then
'            Call CMD_COLLEGAMENTI.SetMenuItemEnabled("Key_Disponibilità", False)
'            Call CMD_COLLEGAMENTI.SetMenuItemEnabled("Key_CicloLavorazione", False)
'            CMD_COLLEGAMENTI.AddMenuItem "Dati scorte", "Key_DatiScorteProd"
'        Else
'            CMD_COLLEGAMENTI.AddMenuItem "Dati scorte e produzione", "Key_DatiScorteProd"
'        End If
'    #End If
'
'    CMD_COLLEGAMENTI.AddMenuItem "Articoli clienti", "Key_ArtClienti"
'    CMD_COLLEGAMENTI.AddMenuItem "Articoli fornitori", "Key_ArtFornitori"
'    CMD_COLLEGAMENTI.AddMenuItem "Scheda prezzi di acquisto", "Key_SkPrezziAcq"
'    CMD_COLLEGAMENTI.AddMenuItem "Scheda prezzi di vendita", "Key_SkPrezziVen"
'
'
'    Exit Sub
'
'Err:
'    Set Gcls_Log.vbError = Err
'    Set Gcls_Log.ADOError = Gcon_Connect.Errors
'    If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.CaricaPgmCollegati") = 1 Then
'       Unload Me
'    Else
'       Resume Next
'    End If
'End Sub

'Private Sub InvocaAnagArticoli()
'Dim AnagArt_Interface As Cinterface
'Dim art As String
'
'    On Error GoTo Err
'
'    Set Cls_ConnectMagazzino.ActiveInterface = ActiveInterface
'    Cls_ConnectMagazzino.Left = 50
'    Cls_ConnectMagazzino.Top = 1000
'    Call Cls_ConnectMagazzino.ArticoloAnagrafica(RTrimN(TXT_CODART.Text))
'    ActiveInterface.IsActive = True
'    Set Cls_ConnectMagazzino.ActiveInterface = Nothing
'    Set ActiveInterface.ClsGlobal.ActiveInterface = ActiveInterface
'    Set ActiveInterface.ActiveFrame = FME_CCS_SKPROD
'    SyncNavigator
'    ActiveInterface.ActiveNavigator.InitializeScript
'
'    Exit Sub
'
'Err:
'    Set Gcls_Log.vbError = Err
'    Set Gcls_Log.ADOError = Gcon_Connect.Errors
'    If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.InvocaAnagArticoli") = 1 Then
'       Unload Me
'    Else
'       Resume Next
'    End If
'End Sub
'
'Private Sub InvocaPartitario()
'Dim Partitario_Interface As Cinterface
'
'    On Error GoTo Err
'
'    Set Cls_ConnectMagazzino.ActiveInterface = ActiveInterface
'    Cls_ConnectMagazzino.Left = 50
'    Cls_ConnectMagazzino.Top = 1000
'    Call Cls_ConnectMagazzino.InterrogazionePartitari(RTrimN(TXT_CODART.Text), RTrimN(TXT_OPZIONE.Text))
'    ActiveInterface.IsActive = True
'    Set Cls_ConnectMagazzino.ActiveInterface = Nothing
'    Set ActiveInterface.ClsGlobal.ActiveInterface = ActiveInterface
'    Set ActiveInterface.ActiveFrame = FME_CCS_SKPROD
'    SyncNavigator
'    ActiveInterface.ActiveNavigator.InitializeScript
'
'    Exit Sub
'
'Err:
'    Set Gcls_Log.vbError = Err
'    Set Gcls_Log.ADOError = Gcon_Connect.Errors
'    If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.InvocaPartitario") = 1 Then
'       Unload Me
'    Else
'       Resume Next
'    End If
'End Sub
'
'Private Sub InvocaDispoProd()
'Dim DispoProd_Interface As Cinterface
'
'    On Error GoTo Err
'
'#If Not GAMMA_SPRINT Then
'
'    Set Pcls_DispoProd = New CLSPD_CCS_ESPLGIA
'    Set DispoProd_Interface = Pcls_DispoProd
'
'    ActiveInterface.IsActive = False
'    Set ActiveInterface.ClsGlobal.ActiveInterface = Pcls_DispoProd
'    ActiveInterface.ClsGlobal.ActiveInterface.IsActive = True
'
'    Set ActiveInterface.ClsGlobal.CallInterface = DispoProd_Interface
'    DispoProd_Interface.IsCalled = True
'
'    Pcls_DispoProd.CodiceArticolo = RTrimN(TXT_CODART.Text)
'    Pcls_DispoProd.Opzione = RTrimN(TXT_OPZIONE.Text)
'    ActiveInterface.ClsGlobal.ExecDll False, "PDUO_CCS_ESPLGIA.CLSPD_CCS_ESPLGIA", True, tsInsert, Normale, 50, 1000
'
'    Set ActiveInterface.ClsGlobal.ActiveInterface = Nothing
'    Set DispoProd_Interface = Nothing
'    Set Pcls_DispoProd = Nothing
'    ActiveInterface.IsActive = True
'    Set ActiveInterface.ClsGlobal.ActiveInterface = ActiveInterface
'    Set ActiveInterface.ClsGlobal.Gcls_VoceMenu = ActiveInterface.ClsVoceMenu
'
'    Set ActiveInterface.ActiveFrame = FME_CCS_SKPROD
'    SyncNavigator
'    ActiveInterface.ActiveNavigator.InitializeScript
'
'#End If
'
'    Exit Sub
'
'Err:
'    Set Gcls_Log.vbError = Err
'    Set Gcls_Log.ADOError = Gcon_Connect.Errors
'    If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.InvocaDispoProd") = 1 Then
'       Unload Me
'    Else
'       Resume Next
'    End If
'End Sub

'Private Sub VisDatiScortaProduzione()
'    On Error GoTo Err
'
'    'Richiamo la FORM x visualizzare i dati delle scorte e di produzione
'    Set FrmScortaProd = New FRMMG_SCORTAPROD
'    Set FrmScortaProd.Gcon_Connect = Gcon_Connect
'    Set FrmScortaProd.Gcls_Log = Gcls_Log
'    Set FrmScortaProd.ActiveInterface = ActiveInterface
'    FrmScortaProd.Gstr_Connect = Gstr_Connect
'    FrmScortaProd.Gstr_DittaCorrente = Gstr_DittaCorrente
'    FrmScortaProd.Articolo = RTrimN(TXT_CODART.Text)
'    FrmScortaProd.Variante = RTrimN(TXT_OPZIONE.Text)
'    FrmScortaProd.Descrizione = RTrimN(TXT_DESCART.Text)
'    If ActiveInterface.WindowModal Then
'        FrmScortaProd.Show vbModal
'    Else
'        Me.Hide
'        FrmScortaProd.Show vbModeless
'    End If
'
'    Exit Sub
'
'Err:
'    Set Gcls_Log.vbError = Err
'    Set Gcls_Log.ADOError = Gcon_Connect.Errors
'    If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.VisDatiScortaProduzione") = 1 Then
'       Unload Me
'    Else
'       Resume Next
'    End If
'End Sub

'Private Sub InvocaCicloLavorazione()
'Dim CicloLavorazione_Interface As Cinterface
'Dim RecCiclo As ADODB.Recordset
'Dim sql As String
'
'    On Error GoTo Err
'
'#If Not GAMMA_SPRINT Then
'
'    Set Pcls_CicloLavorazione = New CLSPD_GESCICLI
'    Set CicloLavorazione_Interface = Pcls_CicloLavorazione
'
'    ActiveInterface.IsActive = False
'    Set ActiveInterface.ClsGlobal.ActiveInterface = CicloLavorazione_Interface
'    ActiveInterface.ClsGlobal.ActiveInterface.IsActive = True
'
'    Set ActiveInterface.ClsGlobal.CallInterface = CicloLavorazione_Interface
'    CicloLavorazione_Interface.IsCalled = True
'
'    'Cerco il codice ciclo dell'articolo
'    ' Set RecCiclo = New ADODB.Recordset
'    sql = "SELECT PD18_CODCICLO FROM PD18_ARTPROD  WITH (NOLOCK) " _
'     & "WHERE PD18_DITTA_CG18 = " & Gstr_DittaCorrente _
'     & " AND PD18_CODART_MG66 = '" & RTrimN(TXT_CODART.Text) & "'"
'
'    Set RecCiclo = Gcon_Connect.Execute(sql, , adCmdText)
'
'    If RecCiclo.EOF = False Then
'        Pcls_CicloLavorazione.CodiceCiclo = RTrimN(RecCiclo.Fields("PD18_CODCICLO").Value)
'        Pcls_CicloLavorazione.VersioneCiclo = 0
'    End If
'    Pcls_CicloLavorazione.NonAbilitareEsporta = True
'
'    ActiveInterface.ClsGlobal.ExecDll False, "PDUO_GESCICLI.CLSPD_GESCICLI", True, tsInsert, Normale, 50, 1000
'
'    Set ActiveInterface.ClsGlobal.ActiveInterface = Nothing
'    Set CicloLavorazione_Interface = Nothing
'    Set Pcls_CicloLavorazione = Nothing
'    ActiveInterface.IsActive = True
'    Set ActiveInterface.ClsGlobal.ActiveInterface = ActiveInterface
'    Set ActiveInterface.ClsGlobal.Gcls_VoceMenu = ActiveInterface.ClsVoceMenu
'
'    Set ActiveInterface.ActiveFrame = FME_CCS_SKPROD
'    SyncNavigator
'    ActiveInterface.ActiveNavigator.InitializeScript
'
'    If Not RecCiclo Is Nothing Then
'        Set RecCiclo = Nothing
'    End If
'
'#End If
'
'    Exit Sub
'
'Err:
'    If Not RecCiclo Is Nothing Then
'        Set RecCiclo = Nothing
'    End If
'    Set Gcls_Log.vbError = Err
'    Set Gcls_Log.ADOError = Gcon_Connect.Errors
'    If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.InvocaPartitario") = 1 Then
'       Unload Me
'    Else
'       Resume Next
'    End If
'End Sub
'
'Private Sub InvocaArtClienti()
'Dim ArtClienti_Interface As Cinterface
'
'    On Error GoTo Err
'
'    Set Pcls_ArtClienti = New CLSMG_ARTCLI
'    Set ArtClienti_Interface = Pcls_ArtClienti
'
'    ActiveInterface.IsActive = False
'    Set ActiveInterface.ClsGlobal.ActiveInterface = ArtClienti_Interface
'    ActiveInterface.ClsGlobal.ActiveInterface.IsActive = True
'
'    Set ActiveInterface.ClsGlobal.CallInterface = ArtClienti_Interface
'    ArtClienti_Interface.IsCalled = True
'
'    Pcls_ArtClienti.CodiceArticolo = RTrimN(TXT_CODART.Text)
'
'    ActiveInterface.ClsGlobal.ExecDll False, "MGUO_ARTCLI.CLSMG_ARTCLI", True, tsInsert, Normale, 50, 1000
'
'    Set ActiveInterface.ClsGlobal.ActiveInterface = Nothing
'    Set ArtClienti_Interface = Nothing
'    Set Pcls_ArtClienti = Nothing
'    ActiveInterface.IsActive = True
'    Set ActiveInterface.ClsGlobal.ActiveInterface = ActiveInterface
'    Set ActiveInterface.ClsGlobal.Gcls_VoceMenu = ActiveInterface.ClsVoceMenu
'
'    Set ActiveInterface.ActiveFrame = FME_CCS_SKPROD
'    SyncNavigator
'    ActiveInterface.ActiveNavigator.InitializeScript
'
'    Exit Sub
'
'Err:
'    Set Gcls_Log.vbError = Err
'    Set Gcls_Log.ADOError = Gcon_Connect.Errors
'    If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.InvocaArtClienti") = 1 Then
'       Unload Me
'    Else
'       Resume Next
'    End If
'End Sub
'
'Private Sub InvocaArtFornitori()
'Dim ArtFornitori_Interface As Cinterface
'
'    On Error GoTo Err
'
'    Set Pcls_ArtFornitori = New CLSMG_ARTFOR
'    Set ArtFornitori_Interface = Pcls_ArtFornitori
'
'    ActiveInterface.IsActive = False
'    Set ActiveInterface.ClsGlobal.ActiveInterface = ArtFornitori_Interface
'    ActiveInterface.ClsGlobal.ActiveInterface.IsActive = True
'
'    Set ActiveInterface.ClsGlobal.CallInterface = ArtFornitori_Interface
'    ArtFornitori_Interface.IsCalled = True
'
'    Pcls_ArtFornitori.CodiceArticolo = RTrimN(TXT_CODART.Text)
'
'    ActiveInterface.ClsGlobal.ExecDll False, "MGUO_ARTFOR.CLSMG_ARTFOR", True, tsInsert, Normale, 50, 1000
'
'    Set ActiveInterface.ClsGlobal.ActiveInterface = Nothing
'    Set ArtFornitori_Interface = Nothing
'    Set Pcls_ArtFornitori = Nothing
'    ActiveInterface.IsActive = True
'    Set ActiveInterface.ClsGlobal.ActiveInterface = ActiveInterface
'    Set ActiveInterface.ClsGlobal.Gcls_VoceMenu = ActiveInterface.ClsVoceMenu
'
'    Set ActiveInterface.ActiveFrame = FME_CCS_SKPROD
'    SyncNavigator
'    ActiveInterface.ActiveNavigator.InitializeScript
'
'    Exit Sub
'
'Err:
'    Set Gcls_Log.vbError = Err
'    Set Gcls_Log.ADOError = Gcon_Connect.Errors
'    If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.InvocaArtFornitori") = 1 Then
'       Unload Me
'    Else
'       Resume Next
'    End If
'End Sub

Private Sub Psub_Elabora(CodArt As String, Variante As String)
On Error GoTo Err
    Dim Pbol_KeyValid As Boolean
    Dim Pstr_Sql      As String
    Dim RsContr       As ADODB.Recordset
        
    Pbol_KeyValid = RTrimN(CodArt) <> "" And TXT_CODART.IsValid
    Set GRID_GIACENZE.DataSource = Nothing
    GRID_GIACENZE.ReBind
    If Not (Pbol_KeyValid) Then
        Exit Sub
    End If

'    CMD_DISPO.Enabled = True
'    CMD_IMPCLI.Enabled = True
'    CMD_ORDFOR.Enabled = True
'    CMD_COLLEGAMENTI.Enabled = True
    
'    If ActiveInterface.ClsGlobal.Gcls_DittaCorrente.ClsDatiGest.FlgProduzione = 1 Then
'        CMD_IMPPROD.Enabled = True
'        CMD_ORDPRO.Enabled = True
'    End If
    
'    Pstr_Sql = "SELECT 'A' FROM SYSOBJECTS " _
'             & "WHERE ID = OBJECT_ID(N'CCS_PREVISIONI') "
'    Set RsContr = Gcon_Connect.Execute(Pstr_Sql, , adCmdText)
'    If Not RsContr.EOF Then
'        CMD_PREIMPCLI.Enabled = True
'    End If
'    Set RsContr = Nothing

    Screen.MousePointer = vbHourglass
    ActiveInterface.StatusBar.Panels(2) = "Elaborazione in corso ..."
    
    Pstr_Sql = ""
    Pstr_Sql = _
                    "SELECT MG70_DITTA_CG18  AS DITTA, " & vbCrLf & _
                    "    MG70_CODDEP_MG58 AS DEPOSITO , " & vbCrLf & _
                    "    MG58_DESCRDEP AS DESCR_DEPOSITO , " & vbCrLf & _
                    "    MG70_TIPOQTA     As TIPOQTA , " & vbCrLf
    If ProgressiviProgetto Then
            Pstr_Sql = Pstr_Sql & _
                    "    MG70_QGIACINI - ISNULL(MG7I_QGIACINI, 0) AS QGIACINI , " & vbCrLf & _
                    "    MG70_QGIACATT - ISNULL(MG7I_QGIACATT, 0) AS QGIACATT , " & vbCrLf & _
                    "    MG70_QGIACEFF - ISNULL(MG7I_QGIACEFF, 0) AS QGIACEFF , " & vbCrLf & _
                    "    MG70_QGIACFIS - ISNULL(MG7I_QGIACFIS, 0) AS QGIACFIS , " & vbCrLf & _
                    "    MG70_QDISPONIB - ISNULL(MG7I_QDISPONIB, 0) AS QDISPONIB , " & vbCrLf & _
                    "    MG70_QIMPCLI - ISNULL(MG7I_QIMPCLI, 0) AS QIMPCLI , " & vbCrLf & _
                    "    MG70_QIMPPROD - ISNULL(MG7I_QIMPPROD, 0) AS QIMPPROD , " & vbCrLf & _
                    "    MG70_QIMPCLAVFOR - ISNULL(MG7I_QIMPCLAVFOR, 0) AS QIMPCLAVFOR , " & vbCrLf & _
                    "    MG70_QPREIMPCLI - ISNULL(MG7I_QPREIMPCLI, 0) AS QPREIMPCLI , " & vbCrLf & _
                    "    MG70_QBLOCSPED - ISNULL(MG7I_QBLOCSPED, 0) AS QBLOCSPED , " & vbCrLf & _
                    "    MG70_QDACONTR - ISNULL(MG7I_QDACONTR, 0) AS QDACONTR , " & vbCrLf & _
                    "    MG70_QORDFOR - ISNULL(MG7I_QORDFOR, 0) AS QORDFOR , " & vbCrLf & _
                    "    MG70_QORDPROD - ISNULL(MG7I_QORDPROD, 0) AS QORDPROD , " & vbCrLf & _
                    "    MG70_QPREIMPFOR - ISNULL(MG7I_QPREIMPFOR, 0) AS QPREIMPFOR , " & vbCrLf & _
                    "    MG70_QDAVAL - ISNULL(MG7I_QDAVAL, 0) AS QDAVAL , " & vbCrLf & _
                    "    MG70_QENTCVIS - ISNULL(MG7I_QENTCVIS, 0) AS QENTCVIS , " & vbCrLf & _
                    "    MG70_QENTCRIP - ISNULL(MG7I_QENTCRIP, 0) AS QENTCRIP , " & vbCrLf & _
                    "    MG70_QENCDEP - ISNULL(MG7I_QENCDEP, 0) AS QENCDEP , " & vbCrLf & _
                    "    MG70_QENCNOLO - ISNULL(MG7I_QENCNOLO, 0) AS QENCNOLO , " & vbCrLf & _
                    "    MG70_QUSCCVIS - ISNULL(MG7I_QUSCCVIS, 0) AS QUSCCVIS , " & vbCrLf & _
                    "    MG70_QUSCCRIP - ISNULL(MG7I_QUSCCRIP, 0) AS QUSCCRIP , " & vbCrLf & _
                    "    MG70_QUSCDEP - ISNULL(MG7I_QUSCDEP, 0) AS QUSCDEP , " & vbCrLf & _
                    "    MG70_QUSCNOLO - ISNULL(MG7I_QUSCNOLO, 0) AS QUSCNOLO , " & vbCrLf

                    'Enzo 200703 - Verifica abilitazione
                    If Not ActiveInterface.ClsVoceMenu.IsRevokeFieldClass(PrezzieScontiImportiAcquisto) Then

Pstr_Sql = Pstr_Sql & "    MG70_QCARACQ - ISNULL(MG7I_QCARACQ, 0) AS QCARACQ , " & vbCrLf & _
                    "    MG70_QCARESORCLI - ISNULL(MG7I_QCARESORCLI, 0) AS QCARESORCLI , " & vbCrLf & _
                    "    MG70_QCARPROD - ISNULL(MG7I_QCARPROD, 0) AS QCARPROD , " & vbCrLf & _
                    "    MG70_QCARCLAVCLI - ISNULL(MG7I_QCARCLAVCLI, 0) AS QCARCLAVCLI , " & vbCrLf & _
                    "    MG70_QCARCLAVFOR - ISNULL(MG7I_QCARCLAVFOR, 0) AS QCARCLAVFOR , " & vbCrLf & _
                    "    MG70_QCAROMAG - ISNULL(MG7I_QCAROMAG, 0) AS QCAROMAG , " & vbCrLf & _
                    "    MG70_QCARGENER - ISNULL(MG7I_QCARGENER, 0) AS QCARGENER , " & vbCrLf & _
                    "    MG70_QCARTRASF - ISNULL(MG7I_QCARTRASF, 0) AS QCARTRASF , " & vbCrLf & _
                    "    MG70_QCARSOST - ISNULL(MG7I_QCARSOST, 0) AS QCARSOST , " & vbCrLf & _
                    "    MG70_QCARLIB1 - ISNULL(MG7I_QCARLIB1, 0) AS QCARLIB1 , " & vbCrLf & _
                    "    MG70_QCARLIB2 - ISNULL(MG7I_QCARLIB2, 0) AS QCARLIB2 , " & vbCrLf & _
                    "    MG70_QSCAVEN - ISNULL(MG7I_QSCAVEN, 0) AS QSCAVEN , " & vbCrLf & _
                    "    MG70_QSCASCART - ISNULL(MG7I_QSCASCART, 0) AS QSCASCART , " & vbCrLf & _
                    "    MG70_QSCAOMAGQ - ISNULL(MG7I_QSCAOMAGQ, 0) AS QSCAOMAGQ , " & vbCrLf & _
                    "    MG70_QSCACLAVCLI - ISNULL(MG7I_QSCACLAVCLI, 0) AS QSCACLAVCLI , " & vbCrLf & _
                    "    MG70_QSCACLAVFOR - ISNULL(MG7I_QSCACLAVFOR, 0) AS QSCACLAVFOR , " & vbCrLf & _
                    "    MG70_QSCAPROD - ISNULL(MG7I_QSCAPROD, 0) AS QSCAPROD , " & vbCrLf & _
                    "    MG70_QSCARESOFOR - ISNULL(MG7I_QSCARESOFOR, 0) AS QSCARESOFOR , " & vbCrLf & _
                    "    MG70_QSCAGENER - ISNULL(MG7I_QSCAGENER, 0) AS QSCAGENER , " & vbCrLf & _
                    "    MG70_QSCATRASF - ISNULL(MG7I_QSCATRASF, 0) AS QSCATRASF , " & vbCrLf & _
                    "    MG70_QSCASOST - ISNULL(MG7I_QSCASOST, 0) AS QSCASOST , " & vbCrLf & _
                    "    MG70_QSCALIB1 - ISNULL(MG7I_QSCALIB1, 0) AS QSCALIB1 , " & vbCrLf & _
                    "    MG70_QSCALIB2 - ISNULL(MG7I_QSCALIB2, 0) As QSCALIB2,"
                    End If
        Pstr_Sql = Pstr_Sql & "    ''         AS COD_PROGETTO ," & vbCrLf & _
                    "    ''           AS DESCR_PROGETTO," & vbCrLf
    Else
        Pstr_Sql = Pstr_Sql & " MG70_QGIACINI AS QGIACINI , " & vbCrLf & _
                    "    MG70_QGIACATT AS QGIACATT , " & vbCrLf & _
                    "    MG70_QGIACEFF AS QGIACEFF , " & vbCrLf & _
                    "    MG70_QGIACFIS AS QGIACFIS , " & vbCrLf & _
                    "    MG70_QDISPONIB AS QDISPONIB , " & vbCrLf & _
                    "    MG70_QIMPCLI AS QIMPCLI , " & vbCrLf & _
                    "    MG70_QIMPPROD AS QIMPPROD , " & vbCrLf & _
                    "    MG70_QIMPCLAVFOR AS QIMPCLAVFOR , " & vbCrLf & _
                    "    MG70_QPREIMPCLI AS QPREIMPCLI , " & vbCrLf & _
                    "    MG70_QBLOCSPED AS QBLOCSPED , " & vbCrLf & _
                    "    MG70_QDACONTR AS QDACONTR , " & vbCrLf & _
                    "    MG70_QORDFOR AS QORDFOR , " & vbCrLf & _
                    "    MG70_QORDPROD AS QORDPROD , " & vbCrLf & _
                    "    MG70_QPREIMPFOR AS QPREIMPFOR , " & vbCrLf & _
                    "    MG70_QDAVAL AS QDAVAL , " & vbCrLf & _
                    "    MG70_QENTCVIS AS QENTCVIS , " & vbCrLf & _
                    "    MG70_QENTCRIP AS QENTCRIP , " & vbCrLf & _
                    "    MG70_QENCDEP AS QENCDEP , " & vbCrLf & _
                    "    MG70_QENCNOLO AS QENCNOLO , " & vbCrLf & _
                    "    MG70_QUSCCVIS AS QUSCCVIS , " & vbCrLf & _
                    "    MG70_QUSCCRIP AS QUSCCRIP , " & vbCrLf
        Pstr_Sql = Pstr_Sql & "  MG70_QUSCDEP AS QUSCDEP , " & vbCrLf & _
                    "    MG70_QUSCNOLO AS QUSCNOLO , " & vbCrLf

                    'Enzo 200703 - Verifica abilitazione
                    If Not ActiveInterface.ClsVoceMenu.IsRevokeFieldClass(PrezzieScontiImportiAcquisto) Then

        Pstr_Sql = Pstr_Sql & "    MG70_QCARACQ AS QCARACQ , " & vbCrLf & _
                    "    MG70_QCARESORCLI AS QCARESORCLI , " & vbCrLf & _
                    "    MG70_QCARPROD AS QCARPROD , " & vbCrLf & _
                    "    MG70_QCARCLAVCLI AS QCARCLAVCLI , " & vbCrLf & _
                    "    MG70_QCARCLAVFOR AS QCARCLAVFOR , " & vbCrLf & _
                    "    MG70_QCAROMAG AS QCAROMAG , " & vbCrLf & _
                    "    MG70_QCARGENER AS QCARGENER , " & vbCrLf & _
                    "    MG70_QCARTRASF AS QCARTRASF , " & vbCrLf & _
                    "    MG70_QCARSOST AS QCARSOST , " & vbCrLf & _
                    "    MG70_QCARLIB1 AS QCARLIB1 , " & vbCrLf & _
                    "    MG70_QCARLIB2 AS QCARLIB2 , " & vbCrLf & _
                    "    MG70_QSCAVEN AS QSCAVEN , " & vbCrLf & _
                    "    MG70_QSCASCART AS QSCASCART , " & vbCrLf & _
                    "    MG70_QSCAOMAGQ AS QSCAOMAGQ , " & vbCrLf & _
                    "    MG70_QSCACLAVCLI AS QSCACLAVCLI , " & vbCrLf & _
                    "    MG70_QSCACLAVFOR AS QSCACLAVFOR , " & vbCrLf & _
                    "    MG70_QSCAPROD AS QSCAPROD , " & vbCrLf & _
                    "    MG70_QSCARESOFOR AS QSCARESOFOR , " & vbCrLf & _
                    "    MG70_QSCAGENER AS QSCAGENER , " & vbCrLf & _
                    "    MG70_QSCATRASF AS QSCATRASF , " & vbCrLf & _
                    "    MG70_QSCASOST AS QSCASOST , " & vbCrLf & _
                    "    MG70_QSCALIB1 AS QSCALIB1 , " & vbCrLf & _
                    "    MG70_QSCALIB2 As QSCALIB2," & vbCrLf
                    End If
    End If
    Pstr_Sql = Pstr_Sql & "    MG70_OPZIONE_MG5E As VARIANTE, " & vbCrLf & _
                          "    0 As TIPOREC "
    Pstr_Sql = Pstr_Sql & "FROM MG70_MAGPROQTA  WITH (NOLOCK) "
    Pstr_Sql = Pstr_Sql & "INNER JOIN MG58_DEPOSITI  WITH (NOLOCK) ON "
    Pstr_Sql = Pstr_Sql & "MG58_DITTA_CG18 = MG70_DITTA_CG18 AND "
    Pstr_Sql = Pstr_Sql & "MG58_CODDEP = MG70_CODDEP_MG58 "


'''
''' Escludo le giacenze a progetto
'''
    If ProgressiviProgetto Then
        Pstr_Sql = Pstr_Sql & " LEFT OUTER JOIN ( "
        Pstr_Sql = Pstr_Sql & "         SELECT MG7G_DITTA_CG18, "
        Pstr_Sql = Pstr_Sql & "             MG7G_CODART_MG66, "
        Pstr_Sql = Pstr_Sql & "             MG7G_OPZIONE_MG5E, "
        Pstr_Sql = Pstr_Sql & "             MG7G_CODDEP_MG58, "
        Pstr_Sql = Pstr_Sql & "             MG7G_TIPOQTA, "
        Pstr_Sql = Pstr_Sql & "             MG7G_PROG_MG4F, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QGIACINI), 0) AS MG7I_QGIACINI, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QGIACATT), 0) AS MG7I_QGIACATT, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QGIACEFF), 0) AS MG7I_QGIACEFF, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QGIACFIS), 0) AS MG7I_QGIACFIS, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QDISPONIB), 0) AS MG7I_QDISPONIB, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QIMPCLI), 0) AS MG7I_QIMPCLI, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QIMPPROD), 0) AS MG7I_QIMPPROD, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QIMPCLAVFOR), 0) AS MG7I_QIMPCLAVFOR, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QPREIMPCLI), 0) AS MG7I_QPREIMPCLI, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QBLOCSPED), 0) AS MG7I_QBLOCSPED, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QDACONTR), 0) AS MG7I_QDACONTR, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QORDFOR), 0) AS MG7I_QORDFOR, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QORDPROD), 0) AS MG7I_QORDPROD, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QPREIMPFOR), 0) AS MG7I_QPREIMPFOR, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QDAVAL), 0) AS MG7I_QDAVAL, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QENTCVIS), 0) AS MG7I_QENTCVIS, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QENTCRIP), 0) AS MG7I_QENTCRIP, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QENCDEP), 0) AS MG7I_QENCDEP, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QENCNOLO), 0) AS MG7I_QENCNOLO, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QUSCCVIS), 0) AS MG7I_QUSCCVIS, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QUSCCRIP), 0) AS MG7I_QUSCCRIP, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QUSCDEP), 0) AS MG7I_QUSCDEP, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QUSCNOLO), 0) AS MG7I_QUSCNOLO "
        
        'Enzo 200703 - Verifica abilitazione
        If Not ActiveInterface.ClsVoceMenu.IsRevokeFieldClass(PrezzieScontiImportiAcquisto) Then
        
        Pstr_Sql = Pstr_Sql & "            , ISNULL(SUM(MG7I_QCARACQ), 0) AS MG7I_QCARACQ, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QCARESORCLI), 0) AS MG7I_QCARESORCLI, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QCARPROD), 0) AS MG7I_QCARPROD, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QCARCLAVCLI), 0) AS MG7I_QCARCLAVCLI, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QCARCLAVFOR), 0) AS MG7I_QCARCLAVFOR, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QCAROMAG), 0) AS MG7I_QCAROMAG, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QCARGENER), 0) AS MG7I_QCARGENER, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QCARTRASF), 0) AS MG7I_QCARTRASF, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QCARSOST), 0) AS MG7I_QCARSOST, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QCARLIB1), 0) AS MG7I_QCARLIB1, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QCARLIB2), 0) AS MG7I_QCARLIB2, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QSCAVEN), 0) AS MG7I_QSCAVEN, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QSCASCART), 0) AS MG7I_QSCASCART, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QSCAOMAGQ), 0) AS MG7I_QSCAOMAGQ, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QSCACLAVCLI), 0) AS MG7I_QSCACLAVCLI, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QSCACLAVFOR), 0) AS MG7I_QSCACLAVFOR, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QSCAPROD), 0) AS MG7I_QSCAPROD, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QSCARESOFOR), 0) AS MG7I_QSCARESOFOR, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QSCAGENER), 0) AS MG7I_QSCAGENER, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QSCATRASF), 0) AS MG7I_QSCATRASF, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QSCASOST), 0) AS MG7I_QSCASOST, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QSCALIB1), 0) AS MG7I_QSCALIB1, "
        Pstr_Sql = Pstr_Sql & "             ISNULL(SUM(MG7I_QSCALIB2), 0) AS MG7I_QSCALIB2 "
        End If
        Pstr_Sql = Pstr_Sql & "         FROM MG7G_PROGQTAVARIREF   WITH (NOLOCK) "
        Pstr_Sql = Pstr_Sql & "         INNER JOIN MG7I_PROGQTAVARI  WITH (NOLOCK) ON MG7G_ID_MG7I = MG7I_ID "
        Pstr_Sql = Pstr_Sql & "         GROUP BY MG7G_DITTA_CG18, "
        Pstr_Sql = Pstr_Sql & "             MG7G_CODART_MG66, "
        Pstr_Sql = Pstr_Sql & "             MG7G_OPZIONE_MG5E, "
        Pstr_Sql = Pstr_Sql & "             MG7G_CODDEP_MG58, "
        Pstr_Sql = Pstr_Sql & "             MG7G_TIPOQTA, "
        Pstr_Sql = Pstr_Sql & "             MG7G_PROG_MG4F) AS MG7G_PROGQTAVARIREF ON MG70_DITTA_CG18 = MG7G_DITTA_CG18  "
        Pstr_Sql = Pstr_Sql & "                     AND MG70_CODART_MG66 = MG7G_CODART_MG66 "
        Pstr_Sql = Pstr_Sql & "                     AND MG70_OPZIONE_MG5E = MG7G_OPZIONE_MG5E "
        Pstr_Sql = Pstr_Sql & "                     AND MG70_CODDEP_MG58 = MG7G_CODDEP_MG58 "
        Pstr_Sql = Pstr_Sql & "                     AND MG70_TIPOQTA = MG7G_TIPOQTA  "
        Pstr_Sql = Pstr_Sql & "                     AND MG7G_PROG_MG4F = " & NumProg & vbCrLf
    End If
    
    'Clausola WHERE
    Pstr_Sql = Pstr_Sql & "WHERE (MG70_DITTA_CG18 = " & Gstr_DittaCorrente
    Pstr_Sql = Pstr_Sql & " OR MG70_DITTA_CG18 IN"
    Pstr_Sql = Pstr_Sql & "( SELECT         MG48_DITTACOL_CG18"
    Pstr_Sql = Pstr_Sql & " FROM MG48_PARMAGAZDC WITH (NOLOCK) "
    Pstr_Sql = Pstr_Sql & " WHERE MG48_DITTA_CG18 = " & Gstr_DittaCorrente & " ) )"
    Pstr_Sql = Pstr_Sql & " AND MG70_CODART_MG66 = '" & RTrimN(CodArt) & "' AND "
    If RTrimN(Variante) > "" Then
        Pstr_Sql = Pstr_Sql & " MG70_OPZIONE_MG5E = " & "'" & RTrimN(Variante) & "'" & " AND "
    End If
    Pstr_Sql = Pstr_Sql & "MG70_TIPOPROG = 1 AND MG70_ANNO = 0 "

    'Filtro sul combo QTA1\QTA2 e sul check solo movimentati
    If CMB_TIPOQTA.Text = 0 Then
        Pstr_Sql = Pstr_Sql & "AND MG70_TIPOQTA = 1 "
    Else
        Pstr_Sql = Pstr_Sql & "AND MG70_TIPOQTA = 2 "
    End If
    
    'Enzo 200707 - Utente del Gruppo Napoli può vedere solamente il magazzino DP
    If ActiveInterface.ClsGlobal.Gcls_UtenteCorrente.Gruppo = "NAPOLI" Then
      Pstr_Sql = Pstr_Sql & " AND MG58_CODDEP IN ('01')"
    Else
      'Enzo 200703 - Filtro per depositi abilitati
      Pstr_Sql = Pstr_Sql & " AND MG58_CODDEP NOT IN ( SELECT FW31_CODTIPO "
      Pstr_Sql = Pstr_Sql & " From FW31_RVKPARAMETRI"
      Pstr_Sql = Pstr_Sql & " WHERE FW31_INDTIPO = 6 "
      ''''''Pstr_Sql = Pstr_Sql & "   AND FW31_GRUPPO = '" & ActiveInterface.ClsGlobal.Gcls_UtenteCorrente.Gruppo & "')"
      Pstr_Sql = Pstr_Sql & "   AND FW31_ID_FW06 = 7)"
    
    End If
    
    
'    If CMB_TIPOQTA.Text = 0 Then
'        If CHK_MOV.Text = 1 Then
'            Pstr_Sql = Pstr_Sql & "AND MG70_TIPOQTA = 1 "
'            If ProgressiviProgetto Then
'                Pstr_Sql = Pstr_Sql & "AND (MG70_QGIACATT - ISNULL(MG7I_QGIACATT,0) <> 0 OR MG70_QGIACEFF - ISNULL(MG7I_QGIACEFF,0) <> 0" _
'                    & "OR MG70_QORDFOR - ISNULL(MG7I_QORDFOR,0) <> 0 OR MG70_QORDPROD - ISNULL(MG7I_QORDPROD,0) <> 0" _
'                    & "OR MG70_QIMPCLI - ISNULL(MG7I_QIMPCLI,0) <> 0 OR MG70_QIMPPROD - ISNULL(MG7I_QIMPPROD,0) <> 0" _
'                    & "OR MG70_QIMPCLAVFOR - ISNULL(MG7I_QIMPCLAVFOR,0) <> 0 OR MG70_QPREIMPCLI - ISNULL(MG7I_QPREIMPCLI,0) <> 0" _
'                    & "OR MG70_QBLOCSPED - ISNULL(MG7I_QBLOCSPED,0) <> 0)"
'            Else
'                Pstr_Sql = Pstr_Sql & " AND (MG70_QGIACATT <> 0 " _
'                    & "OR MG70_QGIACEFF <> 0 OR MG70_QORDFOR <> 0 OR MG70_QORDPROD <> 0 " _
'                    & "OR MG70_QIMPCLI <> 0 OR MG70_QIMPPROD <> 0 OR MG70_QIMPCLAVFOR <> 0 " _
'                    & "OR MG70_QPREIMPCLI <> 0 OR MG70_QBLOCSPED <> 0)"
'            End If
'        Else
'            Pstr_Sql = Pstr_Sql & "AND MG70_TIPOQTA = 1"
'        End If
'    Else
'        If CHK_MOV.Text = 1 Then
'            Pstr_Sql = Pstr_Sql & "AND MG70_TIPOQTA = 2 "
'            If ProgressiviProgetto Then
'                Pstr_Sql = Pstr_Sql & "AND (MG70_QGIACATT - ISNULL(MG7I_QGIACATT,0) <> 0 OR MG70_QGIACEFF - ISNULL(MG7I_QGIACEFF,0) <> 0" _
'                    & "OR MG70_QORDFOR - ISNULL(MG7I_QORDFOR,0) <> 0 OR MG70_QORDPROD - ISNULL(MG7I_QORDPROD,0) <> 0" _
'                    & "OR MG70_QIMPCLI - ISNULL(MG7I_QIMPCLI,0) <> 0 OR MG70_QIMPPROD - ISNULL(MG7I_QIMPPROD,0) <> 0" _
'                    & "OR MG70_QIMPCLAVFOR - ISNULL(MG7I_QIMPCLAVFOR,0) <> 0 OR MG70_QPREIMPCLI - ISNULL(MG7I_QPREIMPCLI,0) <> 0" _
'                    & "OR MG70_QBLOCSPED - ISNULL(MG7I_QBLOCSPED,0) <> 0)"
'            Else
'                Pstr_Sql = Pstr_Sql & " AND (MG70_QGIACATT <> 0 " _
'                    & "OR MG70_QGIACEFF <> 0 OR MG70_QORDFOR <> 0 OR MG70_QORDPROD <> 0 " _
'                    & "OR MG70_QIMPCLI <> 0 OR MG70_QIMPPROD <> 0 OR MG70_QIMPCLAVFOR <> 0 " _
'                    & "OR MG70_QPREIMPCLI <> 0 OR MG70_QBLOCSPED <> 0)"
'            End If
'        Else
'            Pstr_Sql = Pstr_Sql & "AND MG70_TIPOQTA = 2"
'        End If
'    End If

    '''
    ''' Dettaglio progressivi progetto
    '''
    If ProgressiviProgetto Then
    Pstr_Sql = Pstr_Sql & "UNION SELECT MG7G_DITTA_CG18  AS DITTA, " & vbCrLf & _
                    "    MG7G_CODDEP_MG58 AS DEPOSITO , " & vbCrLf & _
                    "    MG58_DESCRDEP AS DESCR_DEPOSITO , " & vbCrLf & _
                    "    MG7G_TIPOQTA     As TIPOQTA , " & vbCrLf & _
                    "    MG7I_QGIACINI    AS QGIACINI , " & vbCrLf & _
                    "    MG7I_QGIACATT    AS QGIACATT , " & vbCrLf & _
                    "    MG7I_QGIACEFF    AS QGIACEFF , " & vbCrLf & _
                    "    MG7I_QGIACFIS    AS QGIACFIS , " & vbCrLf & _
                    "    MG7I_QDISPONIB   AS QDISPONIB , " & vbCrLf & _
                    "    MG7I_QIMPCLI     AS QIMPCLI , " & vbCrLf & _
                    "    MG7I_QIMPPROD    AS QIMPPROD , " & vbCrLf & _
                    "    MG7I_QIMPCLAVFOR AS QIMPCLAVFOR , " & vbCrLf & _
                    "    MG7I_QPREIMPCLI  AS QPREIMPCLI , " & vbCrLf & _
                    "    MG7I_QBLOCSPED   AS QBLOCSPED , " & vbCrLf & _
                    "    MG7I_QDACONTR    AS QDACONTR , " & vbCrLf & _
                    "    MG7I_QORDFOR     AS QORDFOR , " & vbCrLf & _
                    "    MG7I_QORDPROD    AS QORDPROD , " & vbCrLf & _
                    "    MG7I_QPREIMPFOR  AS QPREIMPFOR , " & vbCrLf & _
                    "    MG7I_QDAVAL      AS QDAVAL , " & vbCrLf & _
                    "    MG7I_QENTCVIS    AS QENTCVIS , " & vbCrLf & _
                    "    MG7I_QENTCRIP    AS QENTCRIP , " & vbCrLf & _
                    "    MG7I_QENCDEP     AS QENCDEP , " & vbCrLf & _
                    "    MG7I_QENCNOLO    AS QENCNOLO , " & vbCrLf & _
                    "    MG7I_QUSCCVIS    AS QUSCCVIS , " & vbCrLf & _
                    "    MG7I_QUSCCRIP    AS QUSCCRIP , " & vbCrLf
    Pstr_Sql = Pstr_Sql & "  MG7I_QUSCDEP     AS QUSCDEP , " & vbCrLf & _
                    "    MG7I_QUSCNOLO    AS QUSCNOLO,  " & vbCrLf
                    'Enzo 200703 - Verifica abilitazione
                    If Not ActiveInterface.ClsVoceMenu.IsRevokeFieldClass(PrezzieScontiImportiAcquisto) Then
    Pstr_Sql = Pstr_Sql & "  MG7I_QCARACQ     AS QCARACQ , " & vbCrLf & _
                    "    MG7I_QCARESORCLI AS QCARESORCLI , " & vbCrLf & _
                    "    MG7I_QCARPROD    AS QCARPROD , " & vbCrLf & _
                    "    MG7I_QCARCLAVCLI AS QCARCLAVCLI , " & vbCrLf & _
                    "    MG7I_QCARCLAVFOR AS QCARCLAVFOR , " & vbCrLf & _
                    "    MG7I_QCAROMAG    AS QCAROMAG , " & vbCrLf & _
                    "    MG7I_QCARGENER   AS QCARGENER , " & vbCrLf & _
                    "    MG7I_QCARTRASF   AS QCARTRASF , " & vbCrLf & _
                    "    MG7I_QCARSOST    AS QCARSOST , " & vbCrLf & _
                    "    MG7I_QCARLIB1    AS QCARLIB1 , " & vbCrLf & _
                    "    MG7I_QCARLIB2    AS QCARLIB2 , " & vbCrLf & _
                    "    MG7I_QSCAVEN     AS QSCAVEN , " & vbCrLf & _
                    "    MG7I_QSCASCART   AS QSCASCART , " & vbCrLf & _
                    "    MG7I_QSCAOMAGQ   AS QSCAOMAGQ , " & vbCrLf & _
                    "    MG7I_QSCACLAVCLI AS QSCACLAVCLI , " & vbCrLf & _
                    "    MG7I_QSCACLAVFOR AS QSCACLAVFOR , " & vbCrLf & _
                    "    MG7I_QSCAPROD    AS QSCAPROD , " & vbCrLf & _
                    "    MG7I_QSCARESOFOR AS QSCARESOFOR , " & vbCrLf & _
                    "    MG7I_QSCAGENER   AS QSCAGENER , " & vbCrLf & _
                    "    MG7I_QSCATRASF   AS QSCATRASF , " & vbCrLf & _
                    "    MG7I_QSCASOST    AS QSCASOST , " & vbCrLf & _
                    "    MG7I_QSCALIB1    AS QSCALIB1 , " & vbCrLf & _
                    "    MG7I_QSCALIB2    As QSCALIB2, " & vbCrLf
                    End If
    Pstr_Sql = Pstr_Sql & "    MG7G_CODPROG_PD14 AS COD_PROGETTO ," & vbCrLf & _
                    "    PD68_DESCRPROG   As DESCR_PROGETTO," & vbCrLf & _
                    "    MG7G_OPZIONE_MG5E As VARIANTE, " & vbCrLf & _
                    "    1 As TIPOREC" & vbCrLf & _
                    "FROM MG7G_PROGQTAVARIREF  WITH (NOLOCK) " & vbCrLf & _
                    "INNER JOIN MG7I_PROGQTAVARI WITH (NOLOCK) ON MG7G_ID_MG7I = MG7I_ID" & vbCrLf & _
                    "INNER JOIN MG58_DEPOSITI ON MG58_DITTA_CG18 = MG7G_DITTA_CG18 " & vbCrLf & _
                    "    AND MG58_CODDEP = MG7G_CODDEP_MG58" & vbCrLf & _
                    "INNER JOIN PD68_PROGETTI WITH (NOLOCK) ON MG7G_DITTA_CG18 = PD68_DITTA_CG18" & vbCrLf & _
                    "    AND MG7G_CODPROG_PD14 = PD68_CODPROG" & vbCrLf & _
                    " WHERE (MG7G_DITTA_CG18 = " & Gstr_DittaCorrente & _
                    " OR MG7G_DITTA_CG18 IN" & _
                    "( SELECT         MG48_DITTACOL_CG18" & _
                    " FROM MG48_PARMAGAZDC WITH (NOLOCK) " & _
                    " WHERE MG48_DITTA_CG18 = " & Gstr_DittaCorrente & " ) )" & _
                    "    AND MG7G_CODART_MG66 = '" & RTrimN(CodArt) & "' " & vbCrLf & _
                    "    AND MG7G_PROG_MG4F = " & NumProg & vbCrLf
                    If RTrimN(Variante) > "" Then
                        Pstr_Sql = Pstr_Sql & "    AND MG7G_OPZIONE_MG5E = '" & RTrimN(Variante) & "' " & vbCrLf
                    End If
        'Filtro sul combo QTA1\QTA2 e sul check solo movimentati
        If CMB_TIPOQTA.Text = 0 Then
            Pstr_Sql = Pstr_Sql & "AND MG7G_TIPOQTA = 1"
        Else
            Pstr_Sql = Pstr_Sql & "AND MG7G_TIPOQTA = 2"
        End If
        
'        If CMB_TIPOQTA.Text = 0 Then
'            If CHK_MOV.Text = 1 Then
'                Pstr_Sql = Pstr_Sql & "AND MG7G_TIPOQTA = 1 AND (MG7I_QGIACATT <> 0 " _
'                    & "OR MG7I_QGIACEFF <> 0 OR MG7I_QORDFOR <> 0 OR MG7I_QORDPROD <> 0 " _
'                    & "OR MG7I_QIMPCLI <> 0 OR MG7I_QIMPPROD <> 0 OR MG7I_QIMPCLAVFOR <> 0 " _
'                    & "OR MG7I_QPREIMPCLI <> 0 OR MG7I_QBLOCSPED <> 0)"
'            Else
'                Pstr_Sql = Pstr_Sql & "AND MG7G_TIPOQTA = 1"
'            End If
'        Else
'            If CHK_MOV.Text = 1 Then
'                Pstr_Sql = Pstr_Sql & "AND MG7G_TIPOQTA = 2 AND (MG7I_QGIACATT <> 0 " _
'                    & "OR MG7I_QGIACEFF <> 0 OR MG7I_QORDFOR <> 0 OR MG7I_QORDPROD <> 0 " _
'                    & "OR MG7I_QIMPCLI <> 0 OR MG7I_QIMPPROD <> 0 OR MG7I_QIMPCLAVFOR <> 0 " _
'                    & "OR MG7I_QPREIMPCLI <> 0 OR MG7I_QBLOCSPED <> 0)"
'            Else
'                Pstr_Sql = Pstr_Sql & "AND MG7G_TIPOQTA = 2"
'            End If
'        End If
        Pstr_Sql = Pstr_Sql & "    AND NOT MG7G_CODPROG_PD14 IS NULL "
    End If
    
    'Totali articolo
    Pstr_Sql = Pstr_Sql & "UNION "
    Pstr_Sql = Pstr_Sql & "SELECT "
    Pstr_Sql = Pstr_Sql & "MG70_DITTA_CG18       AS DITTA , "
    Pstr_Sql = Pstr_Sql & "'TOT'                   AS DEPOSITO , "
    Pstr_Sql = Pstr_Sql & "''                    AS DESCR_DEPOSITO , "
    Pstr_Sql = Pstr_Sql & "MG70_TIPOQTA          AS TIPOQTA , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QGIACINI)    AS QGIACINI , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QGIACATT)    AS QGIACATT , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QGIACEFF)    AS QGIACEFF , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QGIACFIS)    AS QGIACFIS , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QDISPONIB)   AS QDISPONIB , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QIMPCLI)     AS QIMPCLI , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QIMPPROD)    AS QIMPPROD , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QIMPCLAVFOR) AS QIMPCLAVFOR , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QPREIMPCLI)  AS QPREIMPCLI , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QBLOCSPED)   AS QBLOCSPED , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QDACONTR)    AS QDACONTR , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QORDFOR)     AS QORDFOR , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QORDPROD)    AS QORDPROD , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QPREIMPFOR)  AS QPREIMPFOR , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QDAVAL)      AS QDAVAL , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QENTCVIS)    AS QENTCVIS , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QENTCRIP)    AS QENTCRIP , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QENCDEP)     AS QENCDEP , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QENCNOLO)    AS QENCNOLO , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QUSCCVIS)    AS QUSCCVIS , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QUSCCRIP)    AS QUSCCRIP , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QUSCDEP)     AS QUSCDEP , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QUSCNOLO)    AS QUSCNOLO , "
    
    'Enzo 200703 - Verifica abilitazione
    If Not ActiveInterface.ClsVoceMenu.IsRevokeFieldClass(PrezzieScontiImportiAcquisto) Then
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QCARACQ)     AS QCARACQ , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QCARESORCLI) AS QCARESORCLI , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QCARPROD)    AS QCARPROD , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QCARCLAVCLI) AS QCARCLAVCLI , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QCARCLAVFOR) AS QCARCLAVFOR , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QCAROMAG)    AS QCAROMAG , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QCARGENER)   AS QCARGENER , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QCARTRASF)   AS QCARTRASF , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QCARSOST)    AS QCARSOST , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QCARLIB1)    AS QCARLIB1 , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QCARLIB2)    AS QCARLIB2 , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QSCAVEN)     AS QSCAVEN , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QSCASCART)   AS QSCASCART , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QSCAOMAGQ)   AS QSCAOMAGQ , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QSCACLAVCLI) AS QSCACLAVCLI , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QSCACLAVFOR) AS QSCACLAVFOR , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QSCAPROD)    AS QSCAPROD , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QSCARESOFOR) AS QSCARESOFOR , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QSCAGENER)   AS QSCAGENER , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QSCATRASF)   AS QSCATRASF , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QSCASOST)    AS QSCASOST , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QSCALIB1)    AS QSCALIB1 , "
    Pstr_Sql = Pstr_Sql & "SUM(MG70_QSCALIB2)    As QSCALIB2, "
    End If

    If ProgressiviProgetto Then
        Pstr_Sql = Pstr_Sql & " '' AS COD_PROGETTO ,"
        Pstr_Sql = Pstr_Sql & " '' AS DESCR_PROGETTO,"
    End If
    Pstr_Sql = Pstr_Sql & " 'TOT.  ARTICOLO' As VARIANTE, "
    Pstr_Sql = Pstr_Sql & "2 As TIPOREC "
    Pstr_Sql = Pstr_Sql & "FROM MG70_MAGPROQTA WITH (NOLOCK) "
    Pstr_Sql = Pstr_Sql & "INNER JOIN MG58_DEPOSITI WITH (NOLOCK) ON "
    Pstr_Sql = Pstr_Sql & "MG58_DITTA_CG18 = MG70_DITTA_CG18 AND "
    Pstr_Sql = Pstr_Sql & "MG58_CODDEP = MG70_CODDEP_MG58 "
    
    'Clausola WHERE
    Pstr_Sql = Pstr_Sql & "WHERE (MG70_DITTA_CG18 = " & Gstr_DittaCorrente
    Pstr_Sql = Pstr_Sql & " OR MG70_DITTA_CG18 IN"
    Pstr_Sql = Pstr_Sql & "( SELECT         MG48_DITTACOL_CG18"
    Pstr_Sql = Pstr_Sql & " FROM MG48_PARMAGAZDC WITH (NOLOCK) "
    Pstr_Sql = Pstr_Sql & " WHERE MG48_DITTA_CG18 = " & Gstr_DittaCorrente & " ) )"
    Pstr_Sql = Pstr_Sql & "AND MG70_CODART_MG66 = '" & RTrimN(CodArt) & "' AND "
    If RTrimN(Variante) > "" Then
        Pstr_Sql = Pstr_Sql & "MG70_OPZIONE_MG5E = " & "'" & RTrimN(Variante) & "'" & " AND "
    End If
    Pstr_Sql = Pstr_Sql & "MG70_TIPOPROG = 1 AND "
    Pstr_Sql = Pstr_Sql & "MG70_ANNO = 0 "
    
    'Filtro sul combo QTA1\QTA2 e sul check solo movimentati
    If CMB_TIPOQTA.Text = 0 Then
        Pstr_Sql = Pstr_Sql & "AND MG70_TIPOQTA = 1 "
    Else
        Pstr_Sql = Pstr_Sql & "AND MG70_TIPOQTA = 2 "
    End If
    
    'Enzo 200707 - Utente del Gruppo Napoli può vedere solamente il magazzino DP
    If ActiveInterface.ClsGlobal.Gcls_UtenteCorrente.Gruppo = "NAPOLI" Then
      Pstr_Sql = Pstr_Sql & " AND MG58_CODDEP IN ('01')"
    Else
      'Enzo 200703 - Filtro per deposito abilitato
      Pstr_Sql = Pstr_Sql & " AND MG58_CODDEP NOT IN ( SELECT FW31_CODTIPO "
      Pstr_Sql = Pstr_Sql & " From FW31_RVKPARAMETRI"
      Pstr_Sql = Pstr_Sql & " WHERE FW31_INDTIPO = 6 "
      Pstr_Sql = Pstr_Sql & "   AND FW31_ID_FW06 = 7 )"
      
      ''''''Pstr_Sql = Pstr_Sql & "   AND FW31_GRUPPO = '" & ActiveInterface.ClsGlobal.Gcls_UtenteCorrente.Gruppo & "')"
    End If
'    If CMB_TIPOQTA.Text = 0 Then
'        If CHK_MOV.Text = 1 Then
'            Pstr_Sql = Pstr_Sql & "AND MG70_TIPOQTA = 1 AND (MG70_QGIACATT <> 0 " _
'                & "OR MG70_QGIACEFF <> 0 OR MG70_QORDFOR <> 0 OR MG70_QORDPROD <> 0 " _
'                & "OR MG70_QIMPCLI <> 0 OR MG70_QIMPPROD <> 0 OR MG70_QIMPCLAVFOR <> 0 " _
'                & "OR MG70_QPREIMPCLI <> 0 OR MG70_QBLOCSPED <> 0)"
'        Else
'            Pstr_Sql = Pstr_Sql & "AND MG70_TIPOQTA = 1"
'        End If
'    Else
'        If CHK_MOV.Text = 1 Then
'            Pstr_Sql = Pstr_Sql & "AND MG70_TIPOQTA = 2 AND (MG70_QGIACATT <> 0 " _
'                & "OR MG70_QGIACEFF <> 0 OR MG70_QORDFOR <> 0 OR MG70_QORDPROD <> 0 " _
'                & "OR MG70_QIMPCLI <> 0 OR MG70_QIMPPROD <> 0 OR MG70_QIMPCLAVFOR <> 0 " _
'                & "OR MG70_QPREIMPCLI <> 0 OR MG70_QBLOCSPED <> 0)"
'        Else
'            Pstr_Sql = Pstr_Sql & "AND MG70_TIPOQTA = 2 "
'        End If
'    End If
    
    'Raggruppamento
    Pstr_Sql = Pstr_Sql & " GROUP BY MG70_DITTA_CG18, MG70_CODART_MG66, MG70_TIPOQTA "
    
    'Ordinamento
    Pstr_Sql = Pstr_Sql & " ORDER BY DITTA, DEPOSITO, MG70_OPZIONE_MG5E ASC"
    'Pstr_Sql = Pstr_Sql & " ORDER BY TIPOREC,MG70_OPZIONE_MG5E,MG70_CODDEP_MG58 "
        
    If Not (Prst_Progressivi Is Nothing) Then
        If Prst_Progressivi.State = adStateOpen Then
            Prst_Progressivi.Close
        End If
        Set Prst_Progressivi = Nothing
    End If
    
    Set Prst_Progressivi = Gcon_Connect.Execute(Pstr_Sql, , adCmdText)
    If Prst_Progressivi.RecordCount > 0 Then
        '
        '   Ordino per descrizione deposito in modo da far trovare la riaga totale per ultimo
        '
        Prst_Progressivi.Sort = "DEPOSITO ASC"
    End If
    
    Set GRID_GIACENZE.DataSource = Prst_Progressivi
    If Prst_Progressivi.RecordCount > 0 Then
       TXT_OPZIONE.Text = RTrimN(Prst_Progressivi("VARIANTE").Value)
    End If
    
    
'    'Enzo 200703 - Carico ultimo prezzo di vendita
'    ' INIZIO ****************************************************************
'
'    Pstr_Sql = " SELECT * FROM LI11_ULTPRACQVEN"
'    Pstr_Sql = Pstr_Sql & " WHERE LI11_DITTA_CG18 = " & Gstr_DittaCorrente
'    Pstr_Sql = Pstr_Sql & " AND LI11_FLGVENACQ = 0 "
'    Pstr_Sql = Pstr_Sql & " AND LI11_CODART_MG66 = '" & RTrimN(CodArt) & "'"
'    Pstr_Sql = Pstr_Sql & " order by LI11_PROG "
'
'    If Not (Grst_RecSet_LI11VEN Is Nothing) Then
'        If Grst_RecSet_LI11VEN.State = adStateOpen Then
'            Grst_RecSet_LI11VEN.Close
'        End If
'        Set Grst_RecSet_LI11VEN = Nothing
'    End If
'
'    Set Grst_RecSet_LI11VEN = Gcon_Connect.Execute(Pstr_Sql, , adCmdText)
'
'    If Not Grst_RecSet_LI11VEN.EOF Then
'      InitializeRecordsetLI11VEN
'    End If
'
'    ' FINE ****************************************************************

    'Enzo 200703 - Carico listini di vendita per articolo
    ' INIZIO ****************************************************************

    Pstr_Sql = " SELECT * FROM LI10_LISTARTIC"
    Pstr_Sql = Pstr_Sql & " WHERE LI10_DITTA_CG18 = " & Gstr_DittaCorrente
    Pstr_Sql = Pstr_Sql & " AND LI10_FLGVENDACQ = 0 "
    Pstr_Sql = Pstr_Sql & " AND LI10_CODART_MG66 = '" & RTrimN(CodArt) & "'"
    Pstr_Sql = Pstr_Sql & " AND     LI10_LISTARTIC.LI10_DATAFINEVAL >= CONVERT(DateTime,'" & Format(Now, "mm/dd/yyyy") & "',101) "
    Pstr_Sql = Pstr_Sql & " AND     LI10_LISTARTIC.LI10_DATAINIZIOVAL <= CONVERT(DateTime,'" & Format(Now, "mm/dd/yyyy") & "',101) "
    Pstr_Sql = Pstr_Sql & " AND     LI10_FLGUSODECOR = 0 "
    Pstr_Sql = Pstr_Sql & " AND LI10_DEPOS_MG58 NOT IN "
    Pstr_Sql = Pstr_Sql & " ( SELECT FW31_CODTIPO "
    Pstr_Sql = Pstr_Sql & "          FROM FW31_RVKPARAMETRI"
    Pstr_Sql = Pstr_Sql & "   WHERE FW31_INDTIPO = 6 "
    Pstr_Sql = Pstr_Sql & "   AND FW31_ID_FW06 = 7"
    
    '''''Pstr_Sql = Pstr_Sql & "     AND FW31_GRUPPO = '" & ActiveInterface.ClsGlobal.Gcls_UtenteCorrente.Gruppo & "'"
    Pstr_Sql = Pstr_Sql & " )"
    Pstr_Sql = Pstr_Sql & " order by LI10_NUMLIST, LI10_DATAINIZIOVAL DESC "

    If Not (Grst_RecSet_LI11VEN Is Nothing) Then
        If Grst_RecSet_LI11VEN.State = adStateOpen Then
            Grst_RecSet_LI11VEN.Close
        End If
        Set Grst_RecSet_LI11VEN = Nothing
    End If

    Set Grst_RecSet_LI11VEN = Gcon_Connect.Execute(Pstr_Sql, , adCmdText)

    Set GRID_LISVEN.DataSource = Grst_RecSet_LI11VEN

    ' FINE ****************************************************************

    'Enzo 200703 - Carico listini di acquisto per articolo
    ' INIZIO ****************************************************************

    Pstr_Sql = " SELECT * FROM LI10_LISTARTIC"
    Pstr_Sql = Pstr_Sql & " WHERE LI10_DITTA_CG18 = " & Gstr_DittaCorrente
    Pstr_Sql = Pstr_Sql & " AND LI10_FLGVENDACQ = 1 "
    Pstr_Sql = Pstr_Sql & " AND LI10_CODART_MG66 = '" & RTrimN(CodArt) & "'"
    Pstr_Sql = Pstr_Sql & " AND     LI10_LISTARTIC.LI10_DATAFINEVAL >= CONVERT(DateTime,'" & Format(Now, "mm/dd/yyyy") & "',101) "
    Pstr_Sql = Pstr_Sql & " AND     LI10_LISTARTIC.LI10_DATAINIZIOVAL <= CONVERT(DateTime,'" & Format(Now, "mm/dd/yyyy") & "',101) "
    Pstr_Sql = Pstr_Sql & " AND LI10_DEPOS_MG58 NOT IN "
    Pstr_Sql = Pstr_Sql & " ( SELECT FW31_CODTIPO "
    Pstr_Sql = Pstr_Sql & "          FROM FW31_RVKPARAMETRI"
    Pstr_Sql = Pstr_Sql & "   WHERE FW31_INDTIPO = 6 "
    '''''''Pstr_Sql = Pstr_Sql & "     AND FW31_GRUPPO = '" & ActiveInterface.ClsGlobal.Gcls_UtenteCorrente.Gruppo & "'"
    Pstr_Sql = Pstr_Sql & "   AND FW31_ID_FW06 = 7"
    Pstr_Sql = Pstr_Sql & " )"
    Pstr_Sql = Pstr_Sql & " order by LI10_NUMLIST "

    If Not (Grst_RecSet_LI11ACQ_TOT Is Nothing) Then
        If Grst_RecSet_LI11ACQ_TOT.State = adStateOpen Then
            Grst_RecSet_LI11ACQ_TOT.Close
        End If
        Set Grst_RecSet_LI11ACQ_TOT = Nothing
    End If

    Set Grst_RecSet_LI11ACQ_TOT = Gcon_Connect.Execute(Pstr_Sql, , adCmdText)
    'Enzo 200703 - Verifica abilitazione
    If ActiveInterface.ClsVoceMenu.IsRevokeFieldClass(PrezzieScontiImportiAcquisto) Then
      GRID_LISACQ_TOT.Visible = False
    Else
      GRID_LISACQ_TOT.Visible = True
      If Not Grst_RecSet_LI11ACQ_TOT.EOF Then
        InitializeRecordsetLI11ACQ_TOT
      Else
        'Pulisci grid
        Set GRID_LISACQ_TOT.DataSource = Grst_RecSet_LI11ACQ_TOT
      End If
    End If

    ' FINE ****************************************************************



'    'Enzo 200703 - Carico ultimo prezzo di acquisto
'    ' INIZIO ****************************************************************
'    Pstr_Sql = " SELECT * FROM LI11_ULTPRACQVEN"
'    Pstr_Sql = Pstr_Sql & " WHERE LI11_DITTA_CG18 = " & Gstr_DittaCorrente
'    Pstr_Sql = Pstr_Sql & " AND LI11_FLGVENACQ = 1 "
'    Pstr_Sql = Pstr_Sql & " AND LI11_CODART_MG66 = '" & RTrimN(CodArt) & "'"
'    Pstr_Sql = Pstr_Sql & " order by LI11_DATAREG DESC "
'
'    If Not (Grst_RecSet_LI11ACQ Is Nothing) Then
'        If Grst_RecSet_LI11ACQ.State = adStateOpen Then
'            Grst_RecSet_LI11ACQ.Close
'        End If
'        Set Grst_RecSet_LI11ACQ = Nothing
'    End If
'
'    Set Grst_RecSet_LI11ACQ = Gcon_Connect.Execute(Pstr_Sql, , adCmdText)
'
'    'Enzo 200703 - Verifica abilitazione
'    If ActiveInterface.ClsVoceMenu.IsRevokeFieldClass(PrezzieScontiImportiAcquisto) Then
'      GRID_LISACQ.Visible = False
'    Else
'      GRID_LISACQ.Visible = True
'      If Not Grst_RecSet_LI11ACQ.EOF Then
'        InitializeRecordsetLI11ACQ
'      End If
'    End If
'    ' FINE ****************************************************************
    
    
'********************************************************************
'Enzo - 20090525 Ultimi due prezzi di acquisto da documenti
'                FOA-DDT, FOA-DDTCLAVCAR
    Pstr_Sql = " SELECT  TOP 2 * "
    Pstr_Sql = Pstr_Sql & " FROM         DO11_DOCTESTATA INNER JOIN"
    Pstr_Sql = Pstr_Sql & "              DO30_DOCCORPO ON DO11_DOCTESTATA.DO11_DITTA_CG18 = DO30_DOCCORPO.DO30_DITTA_CG18 AND "
    Pstr_Sql = Pstr_Sql & "              DO11_DOCTESTATA.DO11_NUMREG_CO99 = DO30_DOCCORPO.DO30_NUMREG_CO99"
    
    Pstr_Sql = Pstr_Sql & " WHERE DO30_CODDEP_MG58 NOT IN "
    Pstr_Sql = Pstr_Sql & " ( SELECT FW31_CODTIPO "
    Pstr_Sql = Pstr_Sql & "          FROM FW31_RVKPARAMETRI"
    Pstr_Sql = Pstr_Sql & "   WHERE FW31_INDTIPO = 6 "
    ''''''''Pstr_Sql = Pstr_Sql & "     AND FW31_GRUPPO = '" & ActiveInterface.ClsGlobal.Gcls_UtenteCorrente.Gruppo & "'"
    Pstr_Sql = Pstr_Sql & "   AND FW31_ID_FW06 = 7"
    Pstr_Sql = Pstr_Sql & " )"
    
    
'    Pstr_Sql = Pstr_Sql & " WHERE DO11_DOCTESTATA.DO11_CODDEP     = '" & RTrimN(Prst_Progressivi("DEPOSITO").Value) & "'"
    Pstr_Sql = Pstr_Sql & "   AND DO30_DOCCORPO.DO30_CODART_MG66  = '" & RTrimN(CodArt) & "'"
    Pstr_Sql = Pstr_Sql & "   AND DO11_DOCTESTATA.DO11_DITTA_CG18 = " & Gstr_DittaCorrente
    Pstr_Sql = Pstr_Sql & "   AND (DO11_DOCTESTATA.DO11_TIPOCF_CG44 = 1)"
    Pstr_Sql = Pstr_Sql & "   AND (DO11_DOCTESTATA.DO11_DOCUM_MG36 IN ('FOA-DDT','FOA-DDTCLAVCAR')) "
    Pstr_Sql = Pstr_Sql & " ORDER BY DO11_DOCTESTATA.DO11_DATAREG DESC "
    
    If Not (Grst_RecSet_LI11ACQ Is Nothing) Then
        If Grst_RecSet_LI11ACQ.State = adStateOpen Then
            Grst_RecSet_LI11ACQ.Close
        End If
        Set Grst_RecSet_LI11ACQ = Nothing
    End If

    Set Grst_RecSet_LI11ACQ = Gcon_Connect.Execute(Pstr_Sql, , adCmdText)
    
    'Enzo 200703 - Verifica abilitazione
    If ActiveInterface.ClsVoceMenu.IsRevokeFieldClass(PrezzieScontiImportiAcquisto) Then
      GRID_LISACQ.Visible = False
    Else
      GRID_LISACQ.Visible = True
      If Not Grst_RecSet_LI11ACQ.EOF Then
        InitializeRecordsetLI11ACQ
      Else
        'Pulisci grid
        Set GRID_LISACQ.DataSource = Grst_RecSet_LI11ACQ
      End If
    End If
    ' FINE ****************************************************************
    
    
    Screen.MousePointer = vbDefault
    ActiveInterface.StatusBar.Panels(2) = "Pronto"
    
    'Disattivo il passaggio dal tasto NUOVO
    ClickNuovo = False
    
    Exit Sub

Err:
  Screen.MousePointer = vbDefault
  ActiveInterface.StatusBar.Panels(2) = "Pronto"
  Set Gcls_Log.vbError = Err
  Set Gcls_Log.ADOError = Gcon_Connect.Errors
  If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.Psub_Elabora") = 1 Then
     Unload Me
  Else
     Resume Next
  End If
End Sub


Private Sub InitializeRecordsetLI11VEN()
    On Error GoTo Err
    
    If Gcls_CalcoloPrezzi Is Nothing Then
        Set Gcls_CalcoloPrezzi = New MGBO_PREZZI.CLSMG_CALCPRNETTO
        Set Gcls_CalcoloPrezzi.ClsDittaCorrente = ActiveInterface.ClsGlobal.Gcls_DittaCorrente
    End If
    
    If Not Grst_RecSet_LI11_appendVEN Is Nothing Then
        If Grst_RecSet_LI11_appendVEN.State = adStateOpen Then
            Grst_RecSet_LI11_appendVEN.Close
        End If
        Set Grst_RecSet_LI11_appendVEN = Nothing
    End If
    Set Grst_RecSet_LI11_appendVEN = New ADODB.Recordset
    
    CreaRecset_Grst_RecSet_LI11VEN
    
    Grst_RecSet_LI11_appendVEN.Open
    
    Trasferisci_in_Grst_RecSet_LI11_AppendVEN
    
    Grst_RecSet_LI11_appendVEN.Filter = "LI11_FLGVENACQ = 0 "
    
    Grst_RecSet_LI11VEN.Filter = "LI11_FLGVENACQ = 0 "
    
    Set GRID_LISVEN.DataSource = Grst_RecSet_LI11_appendVEN
    
    Exit Sub
Err:
    Select Case VisualizzaErrore("InitializeRecordsetLI11VEN")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbCancel
            Resume Next
    End Select
End Sub

Private Sub CreaRecset_Grst_RecSet_LI11VEN()
    On Error GoTo Err
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_DITTA_CG18", adDouble, 30
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_FLGVENACQ", adDecimal, 1
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_INDTIPOLIS", adDecimal, 2
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_CODART_MG66", adBSTR, 30
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_OPZIONE_MG5E", adBSTR, 30
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_DEPOS_MG58", adBSTR, 2
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_PROG", adDecimal, 3
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_CODICE_CG08", adBSTR, 4
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_DATACAMBIO", adDate, 10, adFldMayBeNull
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_CAMBIO", adDouble, 12, adFldMayBeNull
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_PREZZO", adDouble, 20
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_SC1PER", adDouble, 10
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_SC2PER", adDouble, 10
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_SC3PER", adDouble, 10
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_SC4PER", adDouble, 10
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_SC5PER", adDouble, 10
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_SC6PER", adDouble, 10
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_SCIMP", adDouble, 30
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_MAG1PER", adDouble, 10
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_MAG2PER", adDouble, 10
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_MAGIMP", adDouble, 30
    Grst_RecSet_LI11_appendVEN.Fields.Append "PREZZO_NETTO", adDouble, 30
    Grst_RecSet_LI11_appendVEN.Fields.Append "RICARICA", adDouble, 30
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_DATAREG", adDate, 10
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_DATADOC", adDate, 10
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_NUMDOC", adDouble, 10
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_SEZDOC", adBSTR, 2
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_FLGDOCBIS", adBSTR, 2
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_NUMDOCORIG", adBSTR, 10, adFldMayBeNull
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_TIPOCF", adBSTR, 15
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_CODCLFO", adDouble, 10, adFldMayBeNull
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_BVMBASE", adDouble, 30, adFldMayBeNull
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_BVMVAR", adDouble, 30, adFldMayBeNull
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_BVMMOLT", adDouble, 30, adFldMayBeNull
    Grst_RecSet_LI11_appendVEN.Fields.Append "LI11_IDMEDIA_CG99", adDecimal, 20, adFldMayBeNull
    Exit Sub
Err:
    Select Case VisualizzaErrore("CreaRecset_Grst_RecSet_LI11VEN")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbCancel
            Resume Next
    End Select
End Sub

Private Sub Trasferisci_in_Grst_RecSet_LI11_AppendVEN()
    
    Dim strSQLRic  As String
    Dim TotRic     As Double
    Dim ContaRic   As Integer
    
    ContaRic = 0
    TotRic = 0
    
    On Error GoTo Err
    Grst_RecSet_LI11VEN.Filter = adFilterNone
    Grst_RecSet_LI11_appendVEN.Filter = adFilterNone
    If Grst_RecSet_LI11VEN.RecordCount > 0 Then
        Grst_RecSet_LI11VEN.MoveFirst
        While Not Grst_RecSet_LI11VEN.EOF
            Grst_RecSet_LI11_appendVEN.AddNew
            Grst_RecSet_LI11_appendVEN.Fields("LI11_DITTA_CG18").Value = Grst_RecSet_LI11VEN.Fields("LI11_DITTA_CG18").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_FLGVENACQ").Value = Grst_RecSet_LI11VEN.Fields("LI11_FLGVENACQ").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_INDTIPOLIS").Value = Grst_RecSet_LI11VEN.Fields("LI11_INDTIPOLIS").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_CODART_MG66").Value = Grst_RecSet_LI11VEN.Fields("LI11_CODART_MG66").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_OPZIONE_MG5E").Value = Grst_RecSet_LI11VEN.Fields("LI11_OPZIONE_MG5E").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_DEPOS_MG58").Value = Grst_RecSet_LI11VEN.Fields("LI11_DEPOS_MG58").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_PROG").Value = Grst_RecSet_LI11VEN.Fields("LI11_PROG").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_CODICE_CG08").Value = Grst_RecSet_LI11VEN.Fields("LI11_CODICE_CG08").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_DATACAMBIO").Value = Grst_RecSet_LI11VEN.Fields("LI11_DATACAMBIO").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_CAMBIO").Value = Grst_RecSet_LI11VEN.Fields("LI11_CAMBIO").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_PREZZO").Value = Grst_RecSet_LI11VEN.Fields("LI11_PREZZO").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_SC1PER").Value = Grst_RecSet_LI11VEN.Fields("LI11_SC1PER").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_SC2PER").Value = Grst_RecSet_LI11VEN.Fields("LI11_SC2PER").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_SC3PER").Value = Grst_RecSet_LI11VEN.Fields("LI11_SC3PER").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_SC4PER").Value = Grst_RecSet_LI11VEN.Fields("LI11_SC4PER").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_SC5PER").Value = Grst_RecSet_LI11VEN.Fields("LI11_SC5PER").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_SC6PER").Value = Grst_RecSet_LI11VEN.Fields("LI11_SC6PER").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_SCIMP").Value = Grst_RecSet_LI11VEN.Fields("LI11_SCIMP").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_MAG1PER").Value = Grst_RecSet_LI11VEN.Fields("LI11_MAG1PER").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_MAG2PER").Value = Grst_RecSet_LI11VEN.Fields("LI11_MAG2PER").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_MAGIMP").Value = Grst_RecSet_LI11VEN.Fields("LI11_MAGIMP").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_DATAREG").Value = Grst_RecSet_LI11VEN.Fields("LI11_DATAREG").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_DATADOC").Value = Grst_RecSet_LI11VEN.Fields("LI11_DATADOC").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_NUMDOC").Value = Grst_RecSet_LI11VEN.Fields("LI11_NUMDOC").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_SEZDOC").Value = Grst_RecSet_LI11VEN.Fields("LI11_SEZDOC").Value
            Select Case CDecN(Grst_RecSet_LI11VEN.Fields("LI11_FLGDOCBIS").Value)
            Case 0
                Grst_RecSet_LI11_appendVEN.Fields("LI11_FLGDOCBIS").Value = "No"
            Case 1
                Grst_RecSet_LI11_appendVEN.Fields("LI11_FLGDOCBIS").Value = "Si"
            End Select
            Grst_RecSet_LI11_appendVEN.Fields("LI11_NUMDOCORIG").Value = Grst_RecSet_LI11VEN.Fields("LI11_NUMDOCORIG").Value
            Select Case CDecN(Grst_RecSet_LI11VEN.Fields("LI11_TIPOCF").Value)
            Case 0
                Grst_RecSet_LI11_appendVEN.Fields("LI11_TIPOCF").Value = "Cliente"
            Case 1
                Grst_RecSet_LI11_appendVEN.Fields("LI11_TIPOCF").Value = "Fornitore"
            Case 2
                Grst_RecSet_LI11_appendVEN.Fields("LI11_TIPOCF").Value = "Nessuno"
            End Select
            Grst_RecSet_LI11_appendVEN.Fields("LI11_CODCLFO").Value = Grst_RecSet_LI11VEN.Fields("LI11_CODCLFO").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_BVMBASE").Value = Grst_RecSet_LI11VEN.Fields("LI11_BVMBASE").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_BVMVAR").Value = Grst_RecSet_LI11VEN.Fields("LI11_BVMVAR").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_BVMMOLT").Value = Grst_RecSet_LI11VEN.Fields("LI11_BVMMOLT").Value
            Grst_RecSet_LI11_appendVEN.Fields("LI11_IDMEDIA_CG99").Value = Grst_RecSet_LI11VEN.Fields("LI11_IDMEDIA_CG99").Value
                    
            Calcola_PREZZONETTO_LI11VEN
            Gcls_CalcoloPrezzi.CalcolaPrezzoNetto
            Grst_RecSet_LI11_appendVEN.Fields("PREZZO_NETTO") = Gcls_CalcoloPrezzi.PrezzoNetto
            
            
            'Cerco la ricarica dalla LI10
            strSQLRic = " SELECT     *"
            strSQLRic = strSQLRic & " FROM       LI10_LISTARTIC "
            strSQLRic = strSQLRic & " WHERE     LI10_LISTARTIC.LI10_DATAFINEVAL >= CONVERT(DateTime,'" & Format(Now, "mm/dd/yyyy") & "',101) "
            strSQLRic = strSQLRic & "   AND     LI10_LISTARTIC.LI10_DATAINIZIOVAL <= CONVERT(DateTime,'" & Format(Now, "mm/dd/yyyy") & "',101) "
            strSQLRic = strSQLRic & "   AND     LI10_LISTARTIC.LI10_DITTA_CG18 = " & Grst_RecSet_LI11_appendVEN.Fields("LI11_DITTA_CG18").Value
            strSQLRic = strSQLRic & "   AND     LI10_LISTARTIC.LI10_FLGVENDACQ = " & Grst_RecSet_LI11_appendVEN.Fields("LI11_FLGVENACQ").Value
            strSQLRic = strSQLRic & "   AND     LI10_LISTARTIC.LI10_INDTIPOLIS = " & Grst_RecSet_LI11_appendVEN.Fields("LI11_INDTIPOLIS").Value
            strSQLRic = strSQLRic & "   AND     LI10_LISTARTIC.LI10_CODART_MG66 = '" & Grst_RecSet_LI11_appendVEN.Fields("LI11_CODART_MG66").Value & "'"
            strSQLRic = strSQLRic & "   AND     LI10_LISTARTIC.LI10_OPZIONE_MG5E = '" & Grst_RecSet_LI11_appendVEN.Fields("LI11_OPZIONE_MG5E").Value & "'"
            strSQLRic = strSQLRic & "   AND     LI10_LISTARTIC.LI10_DEPOS_MG58 = '" & Grst_RecSet_LI11_appendVEN.Fields("LI11_DEPOS_MG58").Value & "'"
            strSQLRic = strSQLRic & "   AND     LI10_LISTARTIC.LI10_NUMLIST = " & Grst_RecSet_LI11_appendVEN.Fields("LI11_PROG").Value
            
            Set RecDatiAppoggio = Gcon_Connect.Execute(strSQLRic, , adCmdText)
            If RecDatiAppoggio.EOF = False Then
              Grst_RecSet_LI11_appendVEN.Fields("RICARICA") = RecDatiAppoggio.Fields("LI10_PERRICDELTA").Value
              TotRic = TotRic + RecDatiAppoggio.Fields("LI10_PERRICDELTA").Value
              ContaRic = ContaRic + 1
            Else
              Grst_RecSet_LI11_appendVEN.Fields("RICARICA") = 0
            End If
            
            If Not RecDatiAppoggio Is Nothing Then
                Set RecDatiAppoggio = Nothing
            End If
                        
            Grst_RecSet_LI11_appendVEN.UpdateBatch adAffectCurrent
            Grst_RecSet_LI11VEN.MoveNext
        Wend
    End If
    
    If TotRic > 0 And ContaRic > 0 Then
      TXT_RICMEDIA.Text = TotRic / ContaRic
    End If
    
    Grst_RecSet_LI11_appendVEN.Filter = "LI11_FLGVENACQ = 0 "
    
    Exit Sub
Err:
    Select Case VisualizzaErrore("Trasferisci_in_Grst_RecSet_LI11_Append")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbCancel
            Resume Next
    End Select
End Sub

Private Sub Calcola_PREZZONETTO_LI11VEN()
    On Error GoTo Err
    Dim RecSetDB        As ADODB.Recordset
    Dim str_SQL         As String
    
    str_SQL = "SELECT MG66_FLGBASEVAR FROM MG66_ANAGRART WHERE MG66_DITTA_CG18 =" & Gstr_DittaCorrente & " AND MG66_CODART ='" & RTrimN(TXT_CODART.Text) & "'"
    Set RecSetDB = Gcon_Connect.Execute(str_SQL)
    If RecSetDB.RecordCount > 0 Then
        If CDecN(RecSetDB.Fields("MG66_FLGBASEVAR").Value) = 1 Then
            Gcls_CalcoloPrezzi.FlagCalcoloBVM = SiCalcoloBVM
            Gcls_CalcoloPrezzi.Base = CDecN(Grst_RecSet_LI11_appendVEN.Fields("LI11_BVMBASE").Value)
            Gcls_CalcoloPrezzi.Variante = CDecN(Grst_RecSet_LI11_appendVEN.Fields("LI11_BVMVAR").Value)
            Gcls_CalcoloPrezzi.Moltiplicatore = CDecN(Grst_RecSet_LI11_appendVEN.Fields("LI11_BVMMOLT").Value)
        Else
            Gcls_CalcoloPrezzi.FlagCalcoloBVM = NoCalcoloBVM
        End If
    Else
        Gcls_CalcoloPrezzi.FlagCalcoloBVM = NoCalcoloBVM
    End If
    Set RecSetDB = Nothing
    
    Gcls_CalcoloPrezzi.Valuta = RTrimN(Grst_RecSet_LI11_appendVEN.Fields("LI11_CODICE_CG08").Value)
    Gcls_CalcoloPrezzi.PrezzoLordo = RTrimN(Grst_RecSet_LI11_appendVEN.Fields("LI11_PREZZO").Value)
    Gcls_CalcoloPrezzi.Sconto1 = CDecN(Grst_RecSet_LI11_appendVEN.Fields("LI11_SC1PER").Value)
    Gcls_CalcoloPrezzi.Sconto2 = CDecN(Grst_RecSet_LI11_appendVEN.Fields("LI11_SC2PER").Value)
    Gcls_CalcoloPrezzi.Sconto3 = CDecN(Grst_RecSet_LI11_appendVEN.Fields("LI11_SC3PER").Value)
    Gcls_CalcoloPrezzi.Sconto4 = CDecN(Grst_RecSet_LI11_appendVEN.Fields("LI11_SC4PER").Value)
    Gcls_CalcoloPrezzi.Sconto5 = CDecN(Grst_RecSet_LI11_appendVEN.Fields("LI11_SC5PER").Value)
    Gcls_CalcoloPrezzi.Sconto6 = CDecN(Grst_RecSet_LI11_appendVEN.Fields("LI11_SC6PER").Value)
    Gcls_CalcoloPrezzi.ScontoImporto = CDecN(Grst_RecSet_LI11_appendVEN.Fields("LI11_SCIMP").Value)
    Gcls_CalcoloPrezzi.Maggiorazione1 = CDecN(Grst_RecSet_LI11_appendVEN.Fields("LI11_MAG1PER").Value)
    Gcls_CalcoloPrezzi.Maggiorazione2 = CDecN(Grst_RecSet_LI11_appendVEN.Fields("LI11_MAG2PER").Value)
    Gcls_CalcoloPrezzi.MaggiorazioneImporto = CDecN(Grst_RecSet_LI11_appendVEN.Fields("LI11_MAGIMP").Value)
    Exit Sub
Err:
    Select Case VisualizzaErrore("Calcola_PREZZONETTO_LI11VEN")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbCancel
            Resume Next
    End Select
End Sub


Private Sub InitializeRecordsetLI11ACQ()
    On Error GoTo Err
    
    If Gcls_CalcoloPrezzi Is Nothing Then
        Set Gcls_CalcoloPrezzi = New MGBO_PREZZI.CLSMG_CALCPRNETTO
        Set Gcls_CalcoloPrezzi.ClsDittaCorrente = ActiveInterface.ClsGlobal.Gcls_DittaCorrente
    End If
    
    If Not Grst_RecSet_LI11_appendACQ Is Nothing Then
        If Grst_RecSet_LI11_appendACQ.State = adStateOpen Then
            Grst_RecSet_LI11_appendACQ.Close
        End If
        Set Grst_RecSet_LI11_appendACQ = Nothing
    End If
    Set Grst_RecSet_LI11_appendACQ = New ADODB.Recordset
    
    CreaRecset_Grst_RecSet_LI11ACQ
    
    Grst_RecSet_LI11_appendACQ.Open
    
    
'********************************************************************
'Enzo - 20090525 Ultimi due prezzi di acquisto da documenti
'                FOA-DDT, FOA-DDTCLAVCAR
'    Trasferisci_in_Grst_RecSet_LI11_AppendACQ
'
'    Grst_RecSet_LI11_appendACQ.Filter = "LI11_FLGVENACQ = 1 "
'
'    Grst_RecSet_LI11ACQ.Filter = "LI11_FLGVENACQ = 1 "
    
    Trasferisci_in_Grst_RecSet_LI11_AppendACQ_DOC
    
    
    Set GRID_LISACQ.DataSource = Grst_RecSet_LI11_appendACQ
    
    Exit Sub
Err:
    Select Case VisualizzaErrore("InitializeRecordsetLI11ACQ")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbCancel
            Resume Next
    End Select
End Sub

Private Sub CreaRecset_Grst_RecSet_LI11ACQ()
    On Error GoTo Err
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_DITTA_CG18", adDouble, 30
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_FLGVENACQ", adDecimal, 1
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_INDTIPOLIS", adDecimal, 2
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_CODART_MG66", adBSTR, 30
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_OPZIONE_MG5E", adBSTR, 30
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_DEPOS_MG58", adBSTR, 2
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_PROG", adDecimal, 3
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_CODICE_CG08", adBSTR, 4
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_DATACAMBIO", adDate, 10, adFldMayBeNull
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_CAMBIO", adDouble, 12, adFldMayBeNull
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_PREZZO", adDouble, 20
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_SC1PER", adDouble, 10
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_SC2PER", adDouble, 10
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_SC3PER", adDouble, 10
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_SC4PER", adDouble, 10
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_SC5PER", adDouble, 10
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_SC6PER", adDouble, 10
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_SCIMP", adDouble, 30
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_MAG1PER", adDouble, 10
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_MAG2PER", adDouble, 10
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_MAGIMP", adDouble, 30
    Grst_RecSet_LI11_appendACQ.Fields.Append "PREZZO_NETTO", adDouble, 30
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_DATAREG", adDate, 10
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_DATADOC", adDate, 10
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_NUMDOC", adDouble, 10
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_SEZDOC", adBSTR, 2
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_FLGDOCBIS", adBSTR, 2
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_NUMDOCORIG", adBSTR, 10, adFldMayBeNull
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_TIPOCF", adBSTR, 15
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_CODCLFO", adDouble, 10, adFldMayBeNull
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_BVMBASE", adDouble, 30, adFldMayBeNull
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_BVMVAR", adDouble, 30, adFldMayBeNull
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_BVMMOLT", adDouble, 30, adFldMayBeNull
    Grst_RecSet_LI11_appendACQ.Fields.Append "LI11_IDMEDIA_CG99", adDecimal, 20, adFldMayBeNull
    Exit Sub
Err:
    Select Case VisualizzaErrore("CreaRecset_Grst_RecSet_LI11ACQ")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbCancel
            Resume Next
    End Select
End Sub

Private Sub Trasferisci_in_Grst_RecSet_LI11_AppendACQ()
    On Error GoTo Err
    Grst_RecSet_LI11ACQ.Filter = adFilterNone
    Grst_RecSet_LI11_appendACQ.Filter = adFilterNone
    If Grst_RecSet_LI11ACQ.RecordCount > 0 Then
        Grst_RecSet_LI11ACQ.MoveFirst
        While Not Grst_RecSet_LI11ACQ.EOF
            Grst_RecSet_LI11_appendACQ.AddNew
            Grst_RecSet_LI11_appendACQ.Fields("LI11_DITTA_CG18").Value = Grst_RecSet_LI11ACQ.Fields("LI11_DITTA_CG18").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_FLGVENACQ").Value = Grst_RecSet_LI11ACQ.Fields("LI11_FLGVENACQ").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_INDTIPOLIS").Value = Grst_RecSet_LI11ACQ.Fields("LI11_INDTIPOLIS").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_CODART_MG66").Value = Grst_RecSet_LI11ACQ.Fields("LI11_CODART_MG66").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_OPZIONE_MG5E").Value = Grst_RecSet_LI11ACQ.Fields("LI11_OPZIONE_MG5E").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_DEPOS_MG58").Value = Grst_RecSet_LI11ACQ.Fields("LI11_DEPOS_MG58").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_PROG").Value = Grst_RecSet_LI11ACQ.Fields("LI11_PROG").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_CODICE_CG08").Value = Grst_RecSet_LI11ACQ.Fields("LI11_CODICE_CG08").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_DATACAMBIO").Value = Grst_RecSet_LI11ACQ.Fields("LI11_DATACAMBIO").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_CAMBIO").Value = Grst_RecSet_LI11ACQ.Fields("LI11_CAMBIO").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_PREZZO").Value = Grst_RecSet_LI11ACQ.Fields("LI11_PREZZO").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_SC1PER").Value = Grst_RecSet_LI11ACQ.Fields("LI11_SC1PER").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_SC2PER").Value = Grst_RecSet_LI11ACQ.Fields("LI11_SC2PER").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_SC3PER").Value = Grst_RecSet_LI11ACQ.Fields("LI11_SC3PER").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_SC4PER").Value = Grst_RecSet_LI11ACQ.Fields("LI11_SC4PER").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_SC5PER").Value = Grst_RecSet_LI11ACQ.Fields("LI11_SC5PER").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_SC6PER").Value = Grst_RecSet_LI11ACQ.Fields("LI11_SC6PER").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_SCIMP").Value = Grst_RecSet_LI11ACQ.Fields("LI11_SCIMP").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_MAG1PER").Value = Grst_RecSet_LI11ACQ.Fields("LI11_MAG1PER").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_MAG2PER").Value = Grst_RecSet_LI11ACQ.Fields("LI11_MAG2PER").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_MAGIMP").Value = Grst_RecSet_LI11ACQ.Fields("LI11_MAGIMP").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_DATAREG").Value = Grst_RecSet_LI11ACQ.Fields("LI11_DATAREG").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_DATADOC").Value = Grst_RecSet_LI11ACQ.Fields("LI11_DATADOC").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_NUMDOC").Value = Grst_RecSet_LI11ACQ.Fields("LI11_NUMDOC").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_SEZDOC").Value = Grst_RecSet_LI11ACQ.Fields("LI11_SEZDOC").Value
            Select Case CDecN(Grst_RecSet_LI11ACQ.Fields("LI11_FLGDOCBIS").Value)
            Case 0
                Grst_RecSet_LI11_appendACQ.Fields("LI11_FLGDOCBIS").Value = "No"
            Case 1
                Grst_RecSet_LI11_appendACQ.Fields("LI11_FLGDOCBIS").Value = "Si"
            End Select
            Grst_RecSet_LI11_appendACQ.Fields("LI11_NUMDOCORIG").Value = Grst_RecSet_LI11ACQ.Fields("LI11_NUMDOCORIG").Value
            Select Case CDecN(Grst_RecSet_LI11ACQ.Fields("LI11_TIPOCF").Value)
            Case 0
                Grst_RecSet_LI11_appendACQ.Fields("LI11_TIPOCF").Value = "Cliente"
            Case 1
                Grst_RecSet_LI11_appendACQ.Fields("LI11_TIPOCF").Value = "Fornitore"
            Case 2
                Grst_RecSet_LI11_appendACQ.Fields("LI11_TIPOCF").Value = "Nessuno"
            End Select
            Grst_RecSet_LI11_appendACQ.Fields("LI11_CODCLFO").Value = Grst_RecSet_LI11ACQ.Fields("LI11_CODCLFO").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_BVMBASE").Value = Grst_RecSet_LI11ACQ.Fields("LI11_BVMBASE").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_BVMVAR").Value = Grst_RecSet_LI11ACQ.Fields("LI11_BVMVAR").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_BVMMOLT").Value = Grst_RecSet_LI11ACQ.Fields("LI11_BVMMOLT").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_IDMEDIA_CG99").Value = Grst_RecSet_LI11ACQ.Fields("LI11_IDMEDIA_CG99").Value
                    
            Calcola_PREZZONETTO_LI11ACQ
            Gcls_CalcoloPrezzi.CalcolaPrezzoNetto
            Grst_RecSet_LI11_appendACQ.Fields("PREZZO_NETTO") = Gcls_CalcoloPrezzi.PrezzoNetto
            Grst_RecSet_LI11_appendACQ.UpdateBatch adAffectCurrent
            Grst_RecSet_LI11ACQ.MoveNext
        Wend
    End If
    
    Grst_RecSet_LI11_appendACQ.Filter = "LI11_FLGVENACQ = 0 "
    
    Exit Sub
Err:
    Select Case VisualizzaErrore("Trasferisci_in_Grst_RecSet_LI11_Append")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbCancel
            Resume Next
    End Select
End Sub


Private Sub Trasferisci_in_Grst_RecSet_LI11_AppendACQ_DOC()
    On Error GoTo Err
    Grst_RecSet_LI11ACQ.Filter = adFilterNone
    Grst_RecSet_LI11_appendACQ.Filter = adFilterNone
    If Grst_RecSet_LI11ACQ.RecordCount > 0 Then
        Grst_RecSet_LI11ACQ.MoveFirst
        While Not Grst_RecSet_LI11ACQ.EOF
            Grst_RecSet_LI11_appendACQ.AddNew
            Grst_RecSet_LI11_appendACQ.Fields("LI11_DITTA_CG18").Value = Grst_RecSet_LI11ACQ.Fields("DO30_DITTA_CG18").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_FLGVENACQ").Value = Grst_RecSet_LI11ACQ.Fields("DO11_TIPOCF_CG44").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_INDTIPOLIS").Value = 0
            Grst_RecSet_LI11_appendACQ.Fields("LI11_CODART_MG66").Value = Grst_RecSet_LI11ACQ.Fields("DO30_CODART_MG66").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_OPZIONE_MG5E").Value = Grst_RecSet_LI11ACQ.Fields("DO30_OPZIONE_MG5E").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_DEPOS_MG58").Value = Grst_RecSet_LI11ACQ.Fields("DO30_CODDEP_MG58").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_PROG").Value = 1
            Grst_RecSet_LI11_appendACQ.Fields("LI11_CODICE_CG08").Value = NVL(Grst_RecSet_LI11ACQ.Fields("DO11_VALUTA_CG08").Value)
            Grst_RecSet_LI11_appendACQ.Fields("LI11_DATACAMBIO").Value = NVL(Grst_RecSet_LI11ACQ.Fields("DO11_DATACAMBIO").Value)
            Grst_RecSet_LI11_appendACQ.Fields("LI11_CAMBIO").Value = NVL(Grst_RecSet_LI11ACQ.Fields("DO11_CAMBIO").Value)
            Grst_RecSet_LI11_appendACQ.Fields("LI11_PREZZO").Value = Grst_RecSet_LI11ACQ.Fields("DO30_PREZZO1").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_SC1PER").Value = Grst_RecSet_LI11ACQ.Fields("DO30_SCPER1").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_SC2PER").Value = Grst_RecSet_LI11ACQ.Fields("DO30_SCPER2").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_SC3PER").Value = Grst_RecSet_LI11ACQ.Fields("DO30_SCPER3").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_SC4PER").Value = Grst_RecSet_LI11ACQ.Fields("DO30_SCPER4").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_SC5PER").Value = Grst_RecSet_LI11ACQ.Fields("DO30_SCPER5").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_SC6PER").Value = Grst_RecSet_LI11ACQ.Fields("DO30_SCPER6").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_SCIMP").Value = Grst_RecSet_LI11ACQ.Fields("DO30_SCIMP").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_MAG1PER").Value = Grst_RecSet_LI11ACQ.Fields("DO30_MAGPER1").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_MAG2PER").Value = Grst_RecSet_LI11ACQ.Fields("DO30_MAGPER2").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_MAGIMP").Value = Grst_RecSet_LI11ACQ.Fields("DO30_MAGIMP").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_DATAREG").Value = Grst_RecSet_LI11ACQ.Fields("DO11_DATAREG").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_DATADOC").Value = Grst_RecSet_LI11ACQ.Fields("DO11_DATADOC").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_NUMDOC").Value = Grst_RecSet_LI11ACQ.Fields("DO11_NUMDOC").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_SEZDOC").Value = Grst_RecSet_LI11ACQ.Fields("DO11_SEZDOC").Value
            Select Case CDecN(Grst_RecSet_LI11ACQ.Fields("DO11_FLGDOCBIS").Value)
            Case 0
                Grst_RecSet_LI11_appendACQ.Fields("LI11_FLGDOCBIS").Value = "No"
            Case 1
                Grst_RecSet_LI11_appendACQ.Fields("LI11_FLGDOCBIS").Value = "Si"
            End Select
            Grst_RecSet_LI11_appendACQ.Fields("LI11_NUMDOCORIG").Value = Grst_RecSet_LI11ACQ.Fields("DO11_NUMDOCORIG").Value
            Select Case CDecN(Grst_RecSet_LI11ACQ.Fields("DO11_TIPOCF_CG44").Value)
            Case 0
                Grst_RecSet_LI11_appendACQ.Fields("LI11_TIPOCF").Value = "Cliente"
            Case 1
                Grst_RecSet_LI11_appendACQ.Fields("LI11_TIPOCF").Value = "Fornitore"
            Case 2
                Grst_RecSet_LI11_appendACQ.Fields("LI11_TIPOCF").Value = "Nessuno"
            End Select
            Grst_RecSet_LI11_appendACQ.Fields("LI11_CODCLFO").Value = Grst_RecSet_LI11ACQ.Fields("DO11_CLIFOR_CG44").Value
            Grst_RecSet_LI11_appendACQ.Fields("LI11_BVMBASE").Value = Null
            Grst_RecSet_LI11_appendACQ.Fields("LI11_BVMVAR").Value = Null
            Grst_RecSet_LI11_appendACQ.Fields("LI11_BVMMOLT").Value = Null
            Grst_RecSet_LI11_appendACQ.Fields("LI11_IDMEDIA_CG99").Value = Null
            
            Calcola_PREZZONETTO_LI11ACQ
            Gcls_CalcoloPrezzi.CalcolaPrezzoNetto
            Grst_RecSet_LI11_appendACQ.Fields("PREZZO_NETTO") = Gcls_CalcoloPrezzi.PrezzoNetto
            Grst_RecSet_LI11_appendACQ.UpdateBatch adAffectCurrent
            Grst_RecSet_LI11ACQ.MoveNext
        Wend
    End If
    
'    Grst_RecSet_LI11_appendACQ.Filter = "LI11_FLGVENACQ = 0 "
    
    Exit Sub
Err:
    Select Case VisualizzaErrore("Trasferisci_in_Grst_RecSet_LI11_Append")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbCancel
            Resume Next
    End Select
End Sub

Private Sub Calcola_PREZZONETTO_LI11ACQ()
    On Error GoTo Err
    Dim RecSetDB        As ADODB.Recordset
    Dim str_SQL         As String
    
    str_SQL = "SELECT MG66_FLGBASEVAR FROM MG66_ANAGRART WHERE MG66_DITTA_CG18 =" & Gstr_DittaCorrente & " AND MG66_CODART ='" & RTrimN(TXT_CODART.Text) & "'"
    Set RecSetDB = Gcon_Connect.Execute(str_SQL)
    If RecSetDB.RecordCount > 0 Then
        If CDecN(RecSetDB.Fields("MG66_FLGBASEVAR").Value) = 1 Then
            Gcls_CalcoloPrezzi.FlagCalcoloBVM = SiCalcoloBVM
            Gcls_CalcoloPrezzi.Base = CDecN(Grst_RecSet_LI11_appendACQ.Fields("LI11_BVMBASE").Value)
            Gcls_CalcoloPrezzi.Variante = CDecN(Grst_RecSet_LI11_appendACQ.Fields("LI11_BVMVAR").Value)
            Gcls_CalcoloPrezzi.Moltiplicatore = CDecN(Grst_RecSet_LI11_appendACQ.Fields("LI11_BVMMOLT").Value)
        Else
            Gcls_CalcoloPrezzi.FlagCalcoloBVM = NoCalcoloBVM
        End If
    Else
        Gcls_CalcoloPrezzi.FlagCalcoloBVM = NoCalcoloBVM
    End If
    Set RecSetDB = Nothing
    
    Gcls_CalcoloPrezzi.Valuta = RTrimN(Grst_RecSet_LI11_appendACQ.Fields("LI11_CODICE_CG08").Value)
    Gcls_CalcoloPrezzi.PrezzoLordo = RTrimN(Grst_RecSet_LI11_appendACQ.Fields("LI11_PREZZO").Value)
    Gcls_CalcoloPrezzi.Sconto1 = CDecN(Grst_RecSet_LI11_appendACQ.Fields("LI11_SC1PER").Value)
    Gcls_CalcoloPrezzi.Sconto2 = CDecN(Grst_RecSet_LI11_appendACQ.Fields("LI11_SC2PER").Value)
    Gcls_CalcoloPrezzi.Sconto3 = CDecN(Grst_RecSet_LI11_appendACQ.Fields("LI11_SC3PER").Value)
    Gcls_CalcoloPrezzi.Sconto4 = CDecN(Grst_RecSet_LI11_appendACQ.Fields("LI11_SC4PER").Value)
    Gcls_CalcoloPrezzi.Sconto5 = CDecN(Grst_RecSet_LI11_appendACQ.Fields("LI11_SC5PER").Value)
    Gcls_CalcoloPrezzi.Sconto6 = CDecN(Grst_RecSet_LI11_appendACQ.Fields("LI11_SC6PER").Value)
    Gcls_CalcoloPrezzi.ScontoImporto = CDecN(Grst_RecSet_LI11_appendACQ.Fields("LI11_SCIMP").Value)
    Gcls_CalcoloPrezzi.Maggiorazione1 = CDecN(Grst_RecSet_LI11_appendACQ.Fields("LI11_MAG1PER").Value)
    Gcls_CalcoloPrezzi.Maggiorazione2 = CDecN(Grst_RecSet_LI11_appendACQ.Fields("LI11_MAG2PER").Value)
    Gcls_CalcoloPrezzi.MaggiorazioneImporto = CDecN(Grst_RecSet_LI11_appendACQ.Fields("LI11_MAGIMP").Value)
    Exit Sub
Err:
    Select Case VisualizzaErrore("Calcola_PREZZONETTO_LI11ACQ")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbCancel
            Resume Next
    End Select
End Sub

Private Sub InitializeRecordsetLI11ACQ_TOT()
    On Error GoTo Err
    
    If Gcls_CalcoloPrezzi Is Nothing Then
        Set Gcls_CalcoloPrezzi = New MGBO_PREZZI.CLSMG_CALCPRNETTO
        Set Gcls_CalcoloPrezzi.ClsDittaCorrente = ActiveInterface.ClsGlobal.Gcls_DittaCorrente
    End If
    
    If Not Grst_RecSet_LI11_appendACQ_TOT Is Nothing Then
        If Grst_RecSet_LI11_appendACQ_TOT.State = adStateOpen Then
            Grst_RecSet_LI11_appendACQ_TOT.Close
        End If
        Set Grst_RecSet_LI11_appendACQ_TOT = Nothing
    End If
    Set Grst_RecSet_LI11_appendACQ_TOT = New ADODB.Recordset
    
    CreaRecset_Grst_RecSet_LI11ACQ_TOT
    
    Grst_RecSet_LI11_appendACQ_TOT.Open
    
    Trasferisci_in_Grst_RecSet_LI11_AppendACQ_TOT
    
    Grst_RecSet_LI11_appendACQ_TOT.Filter = "LI10_FLGVENDACQ = 1 "
    
    Grst_RecSet_LI11ACQ_TOT.Filter = "LI10_FLGVENDACQ = 1 "
    
    Set GRID_LISACQ_TOT.DataSource = Grst_RecSet_LI11_appendACQ_TOT
    
    Exit Sub
Err:
    Select Case VisualizzaErrore("InitializeRecordsetLI11ACQ_TOT")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbCancel
            Resume Next
    End Select
End Sub

Private Sub CreaRecset_Grst_RecSet_LI11ACQ_TOT()
    On Error GoTo Err
    Grst_RecSet_LI11_appendACQ_TOT.Fields.Append "LI10_DITTA_CG18", adDouble, 30
    Grst_RecSet_LI11_appendACQ_TOT.Fields.Append "LI10_FLGVENDACQ", adDecimal, 1
    Grst_RecSet_LI11_appendACQ_TOT.Fields.Append "LI10_INDTIPOLIS", adDecimal, 2
    Grst_RecSet_LI11_appendACQ_TOT.Fields.Append "LI10_CODART_MG66", adBSTR, 30
    Grst_RecSet_LI11_appendACQ_TOT.Fields.Append "LI10_OPZIONE_MG5E", adBSTR, 30
    Grst_RecSet_LI11_appendACQ_TOT.Fields.Append "LI10_DEPOS_MG58", adBSTR, 2
    Grst_RecSet_LI11_appendACQ_TOT.Fields.Append "LI10_NUMLIST", adDecimal, 3
    Grst_RecSet_LI11_appendACQ_TOT.Fields.Append "LI10_PREZZO", adDouble, 20
    Grst_RecSet_LI11_appendACQ_TOT.Fields.Append "LI10_SC1PER", adDouble, 10
    Grst_RecSet_LI11_appendACQ_TOT.Fields.Append "LI10_SC2PER", adDouble, 10
    Grst_RecSet_LI11_appendACQ_TOT.Fields.Append "LI10_SC3PER", adDouble, 10
    Grst_RecSet_LI11_appendACQ_TOT.Fields.Append "LI10_SC4PER", adDouble, 10
    Grst_RecSet_LI11_appendACQ_TOT.Fields.Append "LI10_SC5PER", adDouble, 10
    Grst_RecSet_LI11_appendACQ_TOT.Fields.Append "LI10_SC6PER", adDouble, 10
    Grst_RecSet_LI11_appendACQ_TOT.Fields.Append "LI10_SCIMP", adDouble, 30
    Grst_RecSet_LI11_appendACQ_TOT.Fields.Append "LI10_MAG1PER", adDouble, 10
    Grst_RecSet_LI11_appendACQ_TOT.Fields.Append "LI10_MAG2PER", adDouble, 10
    Grst_RecSet_LI11_appendACQ_TOT.Fields.Append "LI10_MAGIMP", adDouble, 30
    Grst_RecSet_LI11_appendACQ_TOT.Fields.Append "PREZZO_NETTO", adDouble, 30
    Grst_RecSet_LI11_appendACQ_TOT.Fields.Append "LI10_DATAINIZIOVAL", adDate, 10
    Grst_RecSet_LI11_appendACQ_TOT.Fields.Append "LI10_DATAFINEVAL", adDate, 10
    Grst_RecSet_LI11_appendACQ_TOT.Fields.Append "LI10_BVMBASE", adDouble, 30, adFldMayBeNull
    Grst_RecSet_LI11_appendACQ_TOT.Fields.Append "LI10_BVMVAR", adDouble, 30, adFldMayBeNull
    Grst_RecSet_LI11_appendACQ_TOT.Fields.Append "LI10_BVMMOLT", adDouble, 30, adFldMayBeNull
    Grst_RecSet_LI11_appendACQ_TOT.Fields.Append "LI10_IDMEDIA_CG99", adDecimal, 20, adFldMayBeNull
    Exit Sub
Err:
    Select Case VisualizzaErrore("CreaRecset_Grst_RecSet_LI11ACQ_TOT")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbCancel
            Resume Next
    End Select
End Sub

Private Sub Trasferisci_in_Grst_RecSet_LI11_AppendACQ_TOT()
    On Error GoTo Err
    Grst_RecSet_LI11ACQ_TOT.Filter = adFilterNone
    Grst_RecSet_LI11_appendACQ_TOT.Filter = adFilterNone
    If Grst_RecSet_LI11ACQ_TOT.RecordCount > 0 Then
        Grst_RecSet_LI11ACQ_TOT.MoveFirst
        While Not Grst_RecSet_LI11ACQ_TOT.EOF
            Grst_RecSet_LI11_appendACQ_TOT.AddNew
            Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_DITTA_CG18").Value = Grst_RecSet_LI11ACQ_TOT.Fields("LI10_DITTA_CG18").Value
            Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_FLGVENDACQ").Value = Grst_RecSet_LI11ACQ_TOT.Fields("LI10_FLGVENDACQ").Value
            Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_INDTIPOLIS").Value = Grst_RecSet_LI11ACQ_TOT.Fields("LI10_INDTIPOLIS").Value
            Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_CODART_MG66").Value = Grst_RecSet_LI11ACQ_TOT.Fields("LI10_CODART_MG66").Value
            Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_OPZIONE_MG5E").Value = Grst_RecSet_LI11ACQ_TOT.Fields("LI10_OPZIONE_MG5E").Value
            Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_DEPOS_MG58").Value = Grst_RecSet_LI11ACQ_TOT.Fields("LI10_DEPOS_MG58").Value
            Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_NUMLIST").Value = Grst_RecSet_LI11ACQ_TOT.Fields("LI10_NUMLIST").Value
            Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_PREZZO").Value = Grst_RecSet_LI11ACQ_TOT.Fields("LI10_PREZZO").Value
            Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_SC1PER").Value = Grst_RecSet_LI11ACQ_TOT.Fields("LI10_SC1PER").Value
            Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_SC2PER").Value = Grst_RecSet_LI11ACQ_TOT.Fields("LI10_SC2PER").Value
            Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_SC3PER").Value = Grst_RecSet_LI11ACQ_TOT.Fields("LI10_SC3PER").Value
            Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_SC4PER").Value = Grst_RecSet_LI11ACQ_TOT.Fields("LI10_SC4PER").Value
            Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_SC5PER").Value = Grst_RecSet_LI11ACQ_TOT.Fields("LI10_SC5PER").Value
            Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_SC6PER").Value = Grst_RecSet_LI11ACQ_TOT.Fields("LI10_SC6PER").Value
            Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_SCIMP").Value = Grst_RecSet_LI11ACQ_TOT.Fields("LI10_SCIMP").Value
            Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_MAG1PER").Value = Grst_RecSet_LI11ACQ_TOT.Fields("LI10_MAG1PER").Value
            Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_MAG2PER").Value = Grst_RecSet_LI11ACQ_TOT.Fields("LI10_MAG2PER").Value
            Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_MAGIMP").Value = Grst_RecSet_LI11ACQ_TOT.Fields("LI10_MAGIMP").Value
            Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_DATAINIZIOVAL").Value = Grst_RecSet_LI11ACQ_TOT.Fields("LI10_DATAINIZIOVAL").Value
            Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_DATAFINEVAL").Value = Grst_RecSet_LI11ACQ_TOT.Fields("LI10_DATAFINEVAL").Value
            
            Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_BVMBASE").Value = Grst_RecSet_LI11ACQ_TOT.Fields("LI10_BVMBASE").Value
            Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_BVMVAR").Value = Grst_RecSet_LI11ACQ_TOT.Fields("LI10_BVMVAR").Value
            Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_BVMMOLT").Value = Grst_RecSet_LI11ACQ_TOT.Fields("LI10_BVMMOLT").Value
            Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_IDMEDIA_CG99").Value = Grst_RecSet_LI11ACQ_TOT.Fields("LI10_IDMEDIA_CG99").Value
                    
            Calcola_PREZZONETTO_LI11ACQ_TOT
            Gcls_CalcoloPrezzi.CalcolaPrezzoNetto
            Grst_RecSet_LI11_appendACQ_TOT.Fields("PREZZO_NETTO") = Gcls_CalcoloPrezzi.PrezzoNetto
            Grst_RecSet_LI11_appendACQ_TOT.UpdateBatch adAffectCurrent
            Grst_RecSet_LI11ACQ_TOT.MoveNext
        Wend
    End If
    
    Grst_RecSet_LI11_appendACQ_TOT.Filter = "LI10_FLGVENDACQ = 0 "
    
    Exit Sub
Err:
    Select Case VisualizzaErrore("Trasferisci_in_Grst_RecSet_LI11_Append")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbCancel
            Resume Next
    End Select
End Sub

Private Sub Calcola_PREZZONETTO_LI11ACQ_TOT()
    On Error GoTo Err
    Dim RecSetDB        As ADODB.Recordset
    Dim str_SQL         As String
    
    str_SQL = "SELECT MG66_FLGBASEVAR FROM MG66_ANAGRART WHERE MG66_DITTA_CG18 =" & Gstr_DittaCorrente & " AND MG66_CODART ='" & RTrimN(TXT_CODART.Text) & "'"
    Set RecSetDB = Gcon_Connect.Execute(str_SQL)
    If RecSetDB.RecordCount > 0 Then
        If CDecN(RecSetDB.Fields("MG66_FLGBASEVAR").Value) = 1 Then
            Gcls_CalcoloPrezzi.FlagCalcoloBVM = SiCalcoloBVM
            Gcls_CalcoloPrezzi.Base = CDecN(Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_BVMBASE").Value)
            Gcls_CalcoloPrezzi.Variante = CDecN(Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_BVMVAR").Value)
            Gcls_CalcoloPrezzi.Moltiplicatore = CDecN(Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_BVMMOLT").Value)
        Else
            Gcls_CalcoloPrezzi.FlagCalcoloBVM = NoCalcoloBVM
        End If
    Else
        Gcls_CalcoloPrezzi.FlagCalcoloBVM = NoCalcoloBVM
    End If
    Set RecSetDB = Nothing
    
    Gcls_CalcoloPrezzi.Valuta = "EURO"  'RTrimN(Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_CODICE_CG08").Value)
    Gcls_CalcoloPrezzi.PrezzoLordo = RTrimN(Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_PREZZO").Value)
    Gcls_CalcoloPrezzi.Sconto1 = CDecN(Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_SC1PER").Value)
    Gcls_CalcoloPrezzi.Sconto2 = CDecN(Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_SC2PER").Value)
    Gcls_CalcoloPrezzi.Sconto3 = CDecN(Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_SC3PER").Value)
    Gcls_CalcoloPrezzi.Sconto4 = CDecN(Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_SC4PER").Value)
    Gcls_CalcoloPrezzi.Sconto5 = CDecN(Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_SC5PER").Value)
    Gcls_CalcoloPrezzi.Sconto6 = CDecN(Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_SC6PER").Value)
    Gcls_CalcoloPrezzi.ScontoImporto = CDecN(Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_SCIMP").Value)
    Gcls_CalcoloPrezzi.Maggiorazione1 = CDecN(Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_MAG1PER").Value)
    Gcls_CalcoloPrezzi.Maggiorazione2 = CDecN(Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_MAG2PER").Value)
    Gcls_CalcoloPrezzi.MaggiorazioneImporto = CDecN(Grst_RecSet_LI11_appendACQ_TOT.Fields("LI10_MAGIMP").Value)
    Exit Sub
Err:
    Select Case VisualizzaErrore("Calcola_PREZZONETTO_LI11ACQ_TOT")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbCancel
            Resume Next
    End Select
End Sub


Private Sub FME_CCS_SKPROD_ChangeStatus(ByVal fenm_status As FWBO_VIRTUALFRAME.EnumStatus, ByVal fenm_reason As FWBO_VIRTUALFRAME.EnumReason)

If fenm_status = tsModify Then
    Set GRID_GIACENZE.DataSource = Nothing
    GRID_GIACENZE.ReBind
    If Not ActiveInterface.IsCalled Then
        TXT_OPZIONE.Text = ""
    End If
    Call Psub_Elabora(RTrimN(TXT_CODART.Text), "")
    Call RiempioDati(RTrimN(Grst_SitGiacenze.Fields("MG66_CODART").Value), "")
    CMD_ELABORA.Enabled = False
    TXT_CODART.Enabled = False
    TXT_OPZIONE.Enabled = False
    CMB_TIPOQTA.Enabled = False
End If

End Sub

Private Sub GRID_GIACENZE_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

On Error Resume Next

    If RTrimN(Prst_Progressivi("DEPOSITO").Value) <> "TOT" Then
        TXT_OPZIONE.Text = Prst_Progressivi("VARIANTE").Value
    Else
        TXT_OPZIONE.Text = ""
    End If
    
    
    Dim Pstr_Sql As String
    
    'Enzo 200703 - Seleziono date ultimo carico e scarico per deposito
    'Enzo 200703 - Carico ultimo prezzo di vendita
    ' INIZIO ****************************************************************
    Pstr_Sql = " SELECT TOP 1 MG83_DTULTSCA FROM MG83_PROGDEPOS"
    Pstr_Sql = Pstr_Sql & " WHERE MG83_DITTA_CG18 = " & Gstr_DittaCorrente
    Pstr_Sql = Pstr_Sql & " AND MG83_CODART_MG66 = '" & RTrimN(TXT_CODART.Text) & "'"
    Pstr_Sql = Pstr_Sql & " AND MG83_CODDEP_MG58 = '" & RTrimN(Prst_Progressivi("DEPOSITO").Value) & "'"
    Pstr_Sql = Pstr_Sql & " AND MG83_TIPOPROG = 1 "
    Pstr_Sql = Pstr_Sql & " AND MG83_ANNO = 0 "
'    Pstr_Sql = Pstr_Sql & " order by MG83_ANNO DESC "
    Pstr_Sql = Pstr_Sql & " order by MG83_DTULTSCA DESC "
    
    If Not (Prst_DataCar Is Nothing) Then
        If Prst_DataCar.State = adStateOpen Then
            Prst_DataCar.Close
        End If
        Set Prst_DataCar = Nothing
    End If
    
    Set Prst_DataCar = Gcon_Connect.Execute(Pstr_Sql, , adCmdText)
    
    If Prst_DataCar.EOF = False Then
'Enzo - 200711 Data ultimo carico presa dalle DDT Fornitore
'      TXT_DATAULCA.Text = Prst_DataCar.Fields("MG83_DTULTCAR").Value
      TXT_DATAULSCA.Text = Prst_DataCar.Fields("MG83_DTULTSCA").Value
    Else
'Enzo - 200711 Data ultimo carico presa dalle DDT Fornitore
'      TXT_DATAULCA.Text = ""
      TXT_DATAULSCA.Text = ""
    End If
    
    If Not Prst_DataCar Is Nothing Then
        Set Prst_DataCar = Nothing
    End If
    
    
'********************************************************************
'Enzo - 200711 Data ultimo carico presa dalle DDT Fornitore
    Pstr_Sql = " SELECT  TOP 1 DO11_DOCTESTATA.DO11_DATAREG"
    Pstr_Sql = Pstr_Sql & " FROM         DO11_DOCTESTATA INNER JOIN"
    Pstr_Sql = Pstr_Sql & "                       DO30_DOCCORPO ON DO11_DOCTESTATA.DO11_DITTA_CG18 = DO30_DOCCORPO.DO30_DITTA_CG18 AND "
    Pstr_Sql = Pstr_Sql & "                       DO11_DOCTESTATA.DO11_NUMREG_CO99 = DO30_DOCCORPO.DO30_NUMREG_CO99"
    Pstr_Sql = Pstr_Sql & " WHERE DO11_DOCTESTATA.DO11_CODDEP     = '" & RTrimN(Prst_Progressivi("DEPOSITO").Value) & "'"
    Pstr_Sql = Pstr_Sql & "   AND DO30_DOCCORPO.DO30_CODART_MG66  = '" & RTrimN(TXT_CODART.Text) & "'"
    Pstr_Sql = Pstr_Sql & "   AND DO11_DOCTESTATA.DO11_DITTA_CG18 = " & Gstr_DittaCorrente
    Pstr_Sql = Pstr_Sql & "   AND (DO11_DOCTESTATA.DO11_TIPOCF_CG44 = 1)"
    Pstr_Sql = Pstr_Sql & "   AND (DO11_DOCTESTATA.DO11_STIPODOC IN (1,8,5)) "
    Pstr_Sql = Pstr_Sql & "   AND (DO11_DOCTESTATA.DO11_TIPODOC in ( 1,25)) "
    Pstr_Sql = Pstr_Sql & " ORDER BY DO11_DOCTESTATA.DO11_DATAREG DESC "
    
    If Not (Prst_DataCar Is Nothing) Then
        If Prst_DataCar.State = adStateOpen Then
            Prst_DataCar.Close
        End If
        Set Prst_DataCar = Nothing
    End If
    
    Set Prst_DataCar = Gcon_Connect.Execute(Pstr_Sql, , adCmdText)
    
    If Prst_DataCar.EOF = False Then
'Enzo - 200711 Data ultimo carico presa dalle DDT Fornitore
      TXT_DATAULCA.Text = Prst_DataCar.Fields("DO11_DATAREG").Value
    Else
      TXT_DATAULCA.Text = ""
    End If
    
    If Not Prst_DataCar Is Nothing Then
        Set Prst_DataCar = Nothing
    End If
'Enzo FINE - 200711 Data ultimo carico presa dalle DDT Fornitore
'********************************************************************
    
    
    If ActiveInterface.ClsVoceMenu.IsRevokeFieldClass(PrezzieScontiImportiAcquisto) Then
      TXT_TOTCARICHI.Text = 0
      TXT_TOTSCARICHI.Text = 0
    Else
      TXT_TOTCARICHI.Text = Prst_Progressivi("QCARACQ").Value + Prst_Progressivi("QCARESORCLI").Value + Prst_Progressivi("QCARPROD").Value + Prst_Progressivi("QCARCLAVCLI").Value + Prst_Progressivi("QCARCLAVFOR").Value + Prst_Progressivi("QCAROMAG").Value + Prst_Progressivi("QCARGENER").Value + Prst_Progressivi("QCARTRASF").Value + Prst_Progressivi("QCARSOST").Value + Prst_Progressivi("QCARLIB1").Value + Prst_Progressivi("QCARLIB2").Value
      TXT_TOTSCARICHI.Text = Prst_Progressivi("QSCAVEN").Value + Prst_Progressivi("QSCASCART").Value + Prst_Progressivi("QSCAOMAGQ").Value + Prst_Progressivi("QSCACLAVCLI").Value + Prst_Progressivi("QSCACLAVFOR").Value + Prst_Progressivi("QSCAPROD").Value + Prst_Progressivi("QSCARESOFOR").Value + Prst_Progressivi("QSCAGENER").Value + Prst_Progressivi("QSCATRASF").Value + Prst_Progressivi("QSCASOST").Value + Prst_Progressivi("QSCALIB1").Value + Prst_Progressivi("QSCALIB2").Value
    End If

End Sub

'Private Sub Grst_SitGiacenze_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
''On Error Resume Next
''    If (RTrimN(Grst_SitGiacenze.Fields("MG66_CODART").Value) = "") Or _
''       OnClicLookUp Then
''       Exit Sub
''    End If
''
''    TXT_CODART.Enabled = False
''    Call RiempioDati(RTrimN(Grst_SitGiacenze.Fields("MG66_CODART").Value), "")
''    Call Psub_Elabora(RTrimN(Grst_SitGiacenze.Fields("MG66_CODART").Value), "")
''Err.Clear
'End Sub

Private Sub MDIActiveX1_FormLoad()
    On Error Resume Next
    MDIActiveX1.Move _
        IIf((ActiveInterface.Left = 0), 0, ActiveInterface.Left), _
        IIf((ActiveInterface.Top = 0), 0, ActiveInterface.Top), _
        11830, _
        7075
        
    MDIActiveX1.WindowState = ActiveInterface.WindowState
    MDIActiveX1.MaxButton = False
    MDIActiveX1.OptionSize = False
    Err.Clear
End Sub

Private Sub Form_Activate()
Dim GridColumn          As Column

    On Error GoTo Err
    SyncNavigator
            
    '
    ' Hasanin, 29/05/2006
    '
    If IsLOaded Then
       Exit Sub
    End If
    
'    Set ActiveInterface.ActiveFrame = FME_CCS_SKPROD
    If pbol_alreadyloaded Then
        Exit Sub
    End If
    
    'Esecuzione script personalizzato
'    ExecuteFormEvent ("tsOpen")
    
'    pbol_alreadyloaded = True

    Set Pcls_GridFormat.ActiveInterface = ActiveInterface
    Set Pcls_GridFormat.DataGrid = GRID_GIACENZE
    For Each GridColumn In GRID_GIACENZE.Columns
        If GridColumn.DataField <> "DEPOSITO" And GridColumn.DataField <> "VARIANTE" And GridColumn.DataField <> "DITTA" And _
            GridColumn.DataField <> "DESCR_DEPOSITO" And GridColumn.DataField <> "COD_PROGETTO" And GridColumn.DataField <> "DESCR_PROGETTO" Then
                Set GridColumn.DataFormat = Pstd_Format
        End If
        If GridColumn.DataField = "DEPOSITO" Then
            Set GridColumn.DataFormat = Pstd_FormatDEP
        End If
    Next
    
    ValidateArticolo = True
    ValidateOpzione = True
    
    ' Disabilito la giacenza iniziale e fiscale nella griglia
    GRID_GIACENZE.Columns("Giac.iniziale").Visible = False
    GRID_GIACENZE.Columns("Giac.fiscale").Visible = False
    
    ' Esecuzione layout personalizzato
'    ActiveInterface.ActiveNavigator.ApplyPrsLayout
    
    If ActiveInterface.IsCalled Then
        ' Call TXT_CODART_AfterItem(False)
    Else
        Call CMD_NUOVO_ButtonClick
    End If
    
    If Not ProgressiviProgetto Then
        GRID_GIACENZE.Columns.Item(2).Visible = False
        GRID_GIACENZE.Columns.Item(3).Visible = False
    End If
    
'    FME_CCS_SKPROD.UpdateBatch = False
'    FME_CCS_SKPROD.Status = tsInsert
    Call Psub_Reinizializza
    
    If ActiveInterface.IsCalled Then
        TXT_CODART.Text = RTrimN(ActiveClass.CodiceArticolo)
        If RTrimN(ActiveClass.Opzione) <> "" Then
            DoEvents
            TXT_OPZIONE.Text = ActiveClass.Opzione
        End If
        Old_Articolo = RTrimN(ActiveClass.CodiceArticolo)
        Call Psub_Elabora(RTrimN(TXT_CODART.Text), RTrimN(ActiveClass.Opzione))
        TXT_CODART.Enabled = False
        TXT_OPZIONE.Enabled = False
    End If
    
    IsLOaded = True
Exit Sub

Err:
  Set Gcls_Log.vbError = Err
  Set Gcls_Log.ADOError = Gcon_Connect.Errors
  If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.Form_Active") = 1 Then
     Unload Me
  Else
     Resume Next
  End If
End Sub

Private Sub Form_Load()
On Error GoTo Err
    If ActiveInterface.WindowModal Then
    With Me
        If ActiveInterface.Left = 0 Then
           .Left = 100
        Else
           .Left = ActiveInterface.Left
        End If
        If ActiveInterface.Top = 0 Then
           .Top = 100
        Else
           .Top = ActiveInterface.Top
        End If
        .Width = 11820
        .Height = 7065
    End With
    End If
    
    ' Richiedo l'identificativo della connessione
    Gstr_Connect = ActiveInterface.ClsGlobal.Gcls_LibConnect.GetExtendedProperties
    If Gcon_Connect Is Nothing Then
        Set Gcls_Log = New CLSFW_SrvLog
        Set Gcon_Connect = New ADODB.Connection
        Set Gcon_Connect = Gcls_Connect.Gpr_GetConnect
        With Gcon_Connect
            .ConnectionString = Gstr_Connect
            .Open
        End With
    End If

    Set ActiveInterface.Connection = Gcon_Connect
    
    ' Recupero la ditta
    Gstr_DittaCorrente = ActiveInterface.ClsGlobal.Gcls_DittaCorrente.CodDitta
    
    ' Imposto le classi per le decodifiche e lookup
    Set Pcls_Decode = New MGBO_LOOKUPDECODE.CLSMG_DECODE
    Set Pcls_Decode.ActiveInterface = ActiveInterface
    
    Set Pcls_Lookup = New MGBO_LOOKUPDECODE.CLSMG_LOOKUP
    Set Pcls_Lookup.ActiveInterface = ActiveInterface

    Set Cls_ConnectMagazzino = New MGBO_LOOKUPDECODE.CLSMG_CONNECT
    
#If Not GAMMA_SPRINT Then
    Set Pcls_Connect_Produzione = New PDBO_LOOKUPDECODE.CLSPD_CONNECT
#End If

    ' Carico i combobox
    With CMB_TIPOQTA
        .EraseCombo
        .AddItemData " Quantita' 1", 0
        .AddItemData " Quantita' 2", 1
        .Text = 0
        .AutoOpen = False
    End With
    
    With CMB_TIPOART
        .EraseCombo
        .AddItemData "Standard", 0
        .AddItemData "Standard per progetto", 1
        .AddItemData "Localizzato", 2
        .AddItemData "Pers. per progetto", 3
        .AddItemData "", 4
        .Text = 4
        .AutoOpen = False
    End With
    
    ' Istanzio la classe formattazione grid e data format per la ditta
    Set Pcls_GridFormat = New CLSFW_DataGrid
    
    ClickNuovo = False

    '
    ' Istanzio la classe formattazione grid e data format per la ditta
    '
    Set Pcls_GridFormat = New CLSFW_DataGrid
    Set Pstd_Format = New StdDataFormat
    Set Pstd_FormatDEP = New StdDataFormat
    
    ' Apro i recordset
    Set Grst_SitGiacenze = Gcls_RecSet_SitGiacenze.Gpr_GetADORecord
    Gstr_SQL_SitGiacenze = " SELECT" & _
                           "    MG66_CODART" & _
                           " FROM" & _
                           "    MG66_ANAGRART WITH (NOLOCK) " & _
                           " WHERE" & _
                           "    MG66_DITTA_CG18 = " & Gstr_DittaCorrente
    Grst_SitGiacenze.Open (Gstr_SQL_SitGiacenze & " AND 1=0"), Gcon_Connect
    
    ' Creo una nuova istanza del virtual frame
    Set FME_CCS_SKPROD = New CLSFW_VIRTUALFRAME
    Call FME_CCS_SKPROD.Initialize(ActiveInterface, Gcon_Connect, Grst_SitGiacenze, Gstr_SQL_SitGiacenze, "MG66_CODART")
    FME_CCS_SKPROD.AddControl TXT_CODART
    FME_CCS_SKPROD.AddControl TXT_OPZIONE
    FME_CCS_SKPROD.AddKey TXT_CODART
    FME_CCS_SKPROD.NavigatorSync = False

'    CMD_DISPO.Enabled = False
'    CMD_IMPCLI.Enabled = False
'    CMD_IMPPROD.Enabled = False
'    CMD_ORDFOR.Enabled = False
'    CMD_ORDPRO.Enabled = False
'    CMD_PREIMPCLI.Enabled = False
'    CMD_COLLEGAMENTI.Enabled = False

    OnUnload = False
    '
    ' Hasanin, 29/05/2006
    '
    IsLOaded = False
    
'    Call CaricaPgmCollegati
    
    'Permessi programmi
'    PermAnagrArt = CcsPermessiPrezzi_MENU("MGUO_ARTMAGANUOVI.CLSMG_ARTMAGA")
'    PermPartitario = CcsPermessiPrezzi_MENU("MGUO_INTPART.CLSMG_INTPART")
'    PermCicloLavor = CcsPermessiPrezzi_MENU("PDUO_GESCICLI.CLSPD_GESCICLI")
'    PermDisponibilità = CcsPermessiPrezzi_MENU("PDUO_CCS_ESPLGIA.CLSPD_CCS_ESPLGIA")
'    PermArtClienti = CcsPermessiPrezzi_MENU("MGUO_ARTCLI.CLSMG_ARTCLI")
'    PermArtFornitori = CcsPermessiPrezzi_MENU("MGUO_ARTFOR.CLSMG_ARTFOR")
'    PermSkPrezzi = CcsPermessiPrezzi_MENU("MGUO_CCS_SKPRZART.CLSMG_SCHEDAPRZART")
'
'    If Not PermAnagrArt Then
'        Call CMD_COLLEGAMENTI.SetMenuItemEnabled("Key_Anagrafica", False)
'    End If
'    If Not PermPartitario Then
'        Call CMD_COLLEGAMENTI.SetMenuItemEnabled("Key_Partitario", False)
'    End If
'    If Not PermCicloLavor Then
'        Call CMD_COLLEGAMENTI.SetMenuItemEnabled("Key_CicloLavorazione", False)
'    End If
'    If Not PermDisponibilità Then
'        Call CMD_COLLEGAMENTI.SetMenuItemEnabled("Key_Disponibilità", False)
'    End If
'    If Not PermArtClienti Then
'        Call CMD_COLLEGAMENTI.SetMenuItemEnabled("Key_ArtClienti", False)
'    End If
'    If Not PermArtFornitori Then
'        Call CMD_COLLEGAMENTI.SetMenuItemEnabled("Key_ArtFornitori", False)
'    End If
'    If Not PermSkPrezzi Then
'        Call CMD_COLLEGAMENTI.SetMenuItemEnabled("Key_SkPrezziAcq", False)
'        Call CMD_COLLEGAMENTI.SetMenuItemEnabled("Key_SkPrezziVen", False)
'    End If
'
    Set TXT_CODART.ActiveInterface = ActiveInterface
    Set TXT_CODART.connessione = Gcon_Connect
    TXT_CODART.Ditta = Gstr_DittaCorrente
    
    Set TXT_OPZIONE.ActiveInterface = ActiveInterface
    TXT_OPZIONE.CodiceDitta = Gstr_DittaCorrente
    Set TXT_OPZIONE.GConnect = Gcon_Connect
    TXT_OPZIONE.StringaConnessione = Gstr_Connect
    
    Set TXT_CODART.TxtEditVariante = TXT_OPZIONE
    
'    TXT_OPZIONE.Enabled = True
    
    Call TXT_CODART.MenuEntry("1", "Articoli movimentati", True)
    
'    CMD_DISPO.Visible = False
'    CMD_IMPPROD.Visible = False
'    CMD_ORDPRO.Visible = False
'    TXT_TIPOPROD.Enabled = False
'    CMB_TIPOART.Enabled = False
'    TXT_TIPOPROD.Visible = False
'    CMB_TIPOART.Visible = False
        
    #If Not GAMMA_SPRINT Then
    
    'Enzo 200703 - Nascondo i pulsanti
'        CMD_DISPO.Visible = True
'        CMD_IMPPROD.Visible = True
'        CMD_ORDPRO.Visible = True
'        TXT_TIPOPROD.Visible = True
'        CMB_TIPOART.Visible = True
'        CMB_TIPOART.Enabled = True
        
    #End If
    

'    TMS_RESIZEFORM1.Initialize

    Exit Sub
    
Err:
  Set Gcls_Log.vbError = Err
  Set Gcls_Log.ADOError = Gcon_Connect.Errors
  If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.Form_Load") = 1 Then
     Unload Me
  Else
     Resume Next
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
    SyncNavigator
    ExecuteFormEvent ("tsClose")
    If Not Cancel Then
        Cancel = ActiveInterface.ActiveNavigator.ClsScript.CancelEvent
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim UObject As String

    On Error Resume Next

    OnUnload = True

    If Not ActiveInterface.ActiveNavigator.ClsScript Is Nothing Then
        UObject = ActiveInterface.ClsVoceMenu.Classe
        ActiveInterface.ActiveNavigator.ClsScript.TerminateByUserObject UObject
        Set ActiveInterface.ActiveNavigator.ClsScript = Nothing
    End If
    If Not ActiveInterface.ActiveNavigator.ClsLayout Is Nothing Then
        Set ActiveInterface.ActiveNavigator.ClsLayout = Nothing
    End If
    
    ActiveInterface.ClsGlobal.RemoveCurrentInterface ActiveInterface
    Set ActiveInterface.ClsGlobal.ActiveInterface = Nothing
    Set ActiveInterface.ClsGlobal.CallInterface = Nothing

    Set Pcls_Decode = Nothing
    Set Pcls_Decode.ActiveInterface = Nothing
    Set Pcls_Lookup = Nothing
    Set Pcls_Lookup.ActiveInterface = Nothing
    Set Cls_ConnectMagazzino = Nothing
    
#If Not GAMMA_SPRINT Then
    Set Pcls_Connect_Produzione = Nothing
#End If

    Set cls_datagrid = Nothing

    Set Gcls_Connect = Nothing
    Set Gcls_Log = Nothing
    Gcon_Connect.Close
    Set Gcon_Connect = Nothing

    Set Gcls_Global = Nothing

    Set ActiveInterface.ActiveDll = Nothing
    Set ActiveInterface.ActiveFrame = Nothing
    Set ActiveInterface.Connection = Nothing

    FME_CCS_SKPROD.Terminate
    Set FME_CCS_SKPROD = Nothing

    Set Gcls_RecSet_SitGiacenze = Nothing
    Set Grst_SitGiacenze = Nothing

    Set Pcls_GridFormat = Nothing
    Set Pstd_Format = Nothing
    Set Pstd_FormatDEP = Nothing
        
    'Enzo 200703 - Unload recordset
    Set Gcls_CalcoloPrezzi = Nothing
    If Not (Grst_RecSet_LI11VEN Is Nothing) Then
        If Grst_RecSet_LI11VEN.State = adStateOpen Then
            Grst_RecSet_LI11VEN.Close
        End If
        Set Grst_RecSet_LI11VEN = Nothing
    End If
    
    If Not (Grst_RecSet_LI11_appendVEN Is Nothing) Then
        If Grst_RecSet_LI11_appendVEN.State = adStateOpen Then
            Grst_RecSet_LI11_appendVEN.Close
        End If
        Set Grst_RecSet_LI11_appendVEN = Nothing
    End If
    
    If Not (Grst_RecSet_LI11ACQ Is Nothing) Then
        If Grst_RecSet_LI11ACQ.State = adStateOpen Then
            Grst_RecSet_LI11ACQ.Close
        End If
        Set Grst_RecSet_LI11ACQ = Nothing
    End If
    
    If Not (Grst_RecSet_LI11_appendACQ Is Nothing) Then
        If Grst_RecSet_LI11_appendACQ.State = adStateOpen Then
            Grst_RecSet_LI11_appendACQ.Close
        End If
        Set Grst_RecSet_LI11_appendACQ = Nothing
    End If
    
    If Not (Grst_RecSet_LI11ACQ_TOT Is Nothing) Then
        If Grst_RecSet_LI11ACQ_TOT.State = adStateOpen Then
            Grst_RecSet_LI11ACQ_TOT.Close
        End If
        Set Grst_RecSet_LI11ACQ_TOT = Nothing
    End If
    
    If Not (Grst_RecSet_LI11_appendACQ_TOT Is Nothing) Then
        If Grst_RecSet_LI11_appendACQ_TOT.State = adStateOpen Then
            Grst_RecSet_LI11_appendACQ_TOT.Close
        End If
        Set Grst_RecSet_LI11_appendACQ_TOT = Nothing
    End If
    
    If Not (Prst_DataCar Is Nothing) Then
        If Prst_DataCar.State = adStateOpen Then
            Prst_DataCar.Close
        End If
        Set Prst_DataCar = Nothing
    End If
    
    If Not (RecDatiAppoggio Is Nothing) Then
        If RecDatiAppoggio.State = adStateOpen Then
            RecDatiAppoggio.Close
        End If
        Set RecDatiAppoggio = Nothing
    End If
    
    If Not (Prst_Progressivi Is Nothing) Then
        If Prst_Progressivi.State = adStateOpen Then
            Prst_Progressivi.Close
        End If
        Set Prst_Progressivi = Nothing
    End If

    Set ActiveClass = Nothing
    Set ActiveInterface = Nothing
    
    Err.Clear
End Sub

'Private Sub FrmDispo_CloseFormDispo(Cancel As Boolean)
'    On Error Resume Next
'
'    Set FrmDispo.Gcls_Log = Nothing
'    Set FrmDispo.Gcon_Connect = Nothing
'    Set FrmDispo.ActiveInterface = Nothing
'    Set FrmDispo = Nothing
'    If Not OnUnload Then
'        Me.Show
'    End If
'
'    Err.Clear
'End Sub
'
'Private Sub FrmPreImpCli_CloseFormPreImpCli(Cancel As Boolean)
'    On Error Resume Next
'
'    Set FrmPreImpCli.Gcls_Log = Nothing
'    Set FrmPreImpCli.Gcon_Connect = Nothing
'    Set FrmPreImpCli.ActiveInterface = Nothing
'    Set FrmPreImpCli = Nothing
'    If Not OnUnload Then
'        Me.Show
'    End If
'
'    Err.Clear
'End Sub
'
''Funzione per chiudere il form dei documenti
'Private Sub FrmDocumenti_CloseFormDocumenti(Cancel As Boolean)
'    On Error Resume Next
'
'    Set FrmDocumenti.Gcls_Log = Nothing
'    Set FrmDocumenti.Gcon_Connect = Nothing
'    Set FrmDocumenti.ActiveInterface = Nothing
'    Set FrmDocumenti = Nothing
'    If Not OnUnload Then
'        Me.Show
'    End If
'
'    Err.Clear
'End Sub
'
'Private Sub FrmODL_CloseFormODL(Cancel As Boolean)
'    On Error Resume Next
'
'    Set FrmODL.Gcls_Log = Nothing
'    Set FrmODL.Gcon_Connect = Nothing
'    Set FrmODL.ActiveInterface = Nothing
'    Set FrmODL = Nothing
'    If Not OnUnload Then
'        Me.Show
'    End If
'
'    Err.Clear
'End Sub
'
'Private Sub FrmImprod_CloseFormImprod(Cancel As Boolean)
'    On Error Resume Next
'
'    Set FrmImprod.Gcls_Log = Nothing
'    Set FrmImprod.Gcon_Connect = Nothing
'    Set FrmImprod.ActiveInterface = Nothing
'    Set FrmImprod = Nothing
'    If Not OnUnload Then
'        Me.Show
'    End If
'
'    Err.Clear
'End Sub
'
'Private Sub FrmScortaProd_CloseFormScorteProd(Cancel As Boolean)
'    On Error Resume Next
'
'    Set FrmScortaProd.Gcls_Log = Nothing
'    Set FrmScortaProd.Gcon_Connect = Nothing
'    Set FrmScortaProd.ActiveInterface = Nothing
'    Set FrmScortaProd = Nothing
'    If Not OnUnload Then
'        Me.Show
'    End If
'
'    Err.Clear
'End Sub

Private Sub Pstd_Format_Format(ByVal DataValue As StdFormat.StdDataValue)

        If pvarDecimali > 0 Then
            DataValue = Format(DataValue, "##,##0." & String(pvarDecimali, "0"))
        Else
            DataValue = Format(DataValue, "##,##0")
        End If

End Sub

'Private Sub Pstd_FormatDEP_Format(ByVal DataValue As StdFormat.StdDataValue)
'
'    If DataValue = "z" Then
'        DataValue = "TOT"
'    End If
'
'End Sub
'
'Private Sub TMS_RESIZEFORM1_BeforeAutoInitialize(DisableAutoInitialize As Boolean)
'
'    On Error Resume Next
'
'    DisableAutoInitialize = True
'    TMS_RESIZEFORM1.AddControl GRID_GIACENZE, tsAnchorleft Or tsAnchorRight Or tsAnchorBottom Or tsAnchorTop
''    TMS_RESIZEFORM1.AddControl CMD_COLLEGAMENTI, tsAnchorBottom Or tsAnchorleft
''    TMS_RESIZEFORM1.AddControl CMD_DISPO, tsAnchorBottom Or tsAnchorRight
''    TMS_RESIZEFORM1.AddControl CMD_IMPCLI, tsAnchorBottom Or tsAnchorRight
''    TMS_RESIZEFORM1.AddControl CMD_IMPPROD, tsAnchorBottom Or tsAnchorRight
'    TMS_RESIZEFORM1.AddControl CMD_NUOVO, tsAnchorBottom Or tsAnchorRight
''    TMS_RESIZEFORM1.AddControl CMD_ORDFOR, tsAnchorBottom Or tsAnchorRight
''    TMS_RESIZEFORM1.AddControl CMD_ORDPRO, tsAnchorBottom Or tsAnchorRight
''    TMS_RESIZEFORM1.AddControl CMD_PREIMPCLI, tsAnchorBottom Or tsAnchorRight
'    TMS_RESIZEFORM1.AddControl CMD_ELABORA, tsAnchorBottom Or tsAnchorRight
'    Err.Clear
'
'End Sub

Private Sub TXT_CODART_AfterChange(Cancel As Boolean)

On Error Resume Next

FME_CCS_SKPROD.UpdateBatch = False

End Sub

Private Sub TXT_CODART_GotFocus()

On Error Resume Next

FME_CCS_SKPROD.UpdateBatch = True

End Sub

'Private Sub TXT_CODART_GotFocus()
'
''Dim Old_Art                     As String
''Dim Old_Opzione                 As String
''
''Old_Art = RTrimN(TXT_CODART.Text)
''Old_Opzione = RTrimN(TXT_OPZIONE.Text)
''
''If Not FME_CCS_SKPROD.Status = tsInsert Then
''    FME_CCS_SKPROD.Status = tsInsert
''End If
''
''TXT_CODART.Text = Old_Art
''If Old_Opzione <> "" Then
''    TXT_OPZIONE.Text = Old_Opzione
''End If
'
'End Sub

Private Sub TXT_CODART_KeyButtonMenuPress(Cancel As Boolean, ByVal Pstr_KeyButtonPress As Variant)
    On Error Resume Next

Select Case Pstr_KeyButtonPress
    Case "Kpers1"
        Pint_LookupPers = 1
        FME_CCS_SKPROD.UpdateBatch = True
        'TXT_CODART.CanReturnRecordSet = True
        TXT_CODART.StartLookup
        'TXT_CODART.CanReturnRecordSet = False
        FME_CCS_SKPROD.UpdateBatch = False
        Pint_LookupPers = 0
End Select

'    If Pstr_KeyButtonPress = "Kpers1" Then
'       PbolLookupArticForn = True
'       TXT_CODART.StartLookup
'       PbolLookupArticForn = False
'    Else
'        If Pstr_KeyButtonPress = "Kpers2" Then
'           PbolLookupArticCli = True
'           TXT_CODART.StartLookup
'           PbolLookupArticCli = False
'        End If
'    End If
    
    Err.Clear
End Sub

Private Sub TXT_CODART_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
    Cancel = False
    On Error GoTo Err
    
    Dim Pst_Colonne(0 To 11, 0 To 1) As Variant
    
    If Pint_LookupPers = 0 Then
        OnClicLookUp = True
        Cancel = False
        Pcls_Lookup.ArticoliDiMagazzino
        str_SQL = Pcls_Lookup.StringaSQL
        Arr_Fields = Pcls_Lookup.ArrayFields
        Str_Caption = Pcls_Lookup.Titolo
        Str_Connect = Gstr_Connect
        TXT_CODART.IDLookup = Pcls_Lookup.IDLookup
        OnClicLookUp = False
    Else
        '
        '   Stringa SQL della lookup articoli con giacenza
        '
str_SQL = ""
str_SQL = _
                    "SELECT DISTINCT MG70_CODART_MG66, " & vbCrLf & _
                    "    MG66_CODINTERNO," & vbCrLf & _
                    "    MG66_UM1," & vbCrLf & _
                    "    MG66_FAM_MG53," & vbCrLf & _
                    "    MG53_DESCRFAM," & vbCrLf & _
                    "    MG66_SFAM_MG54," & vbCrLf & _
                    "    MG54_DESCRSFAM," & vbCrLf & _
                    "    MG66_GRUPPO_MG55," & vbCrLf & _
                    "    MG55_DESCRGRUPPO," & vbCrLf & _
                    "    MG66_SGRUPPO_MG56," & vbCrLf & _
                    "    MG56_DESCRSGRUPPO" & vbCrLf & _
                    "FROM MG70_MAGPROQTA  WITH (NOLOCK) " & vbCrLf & _
                    "INNER JOIN MG66_ANAGRART  WITH (NOLOCK) ON MG66_DITTA_CG18 = MG70_DITTA_CG18" & vbCrLf & _
                    "    AND MG66_CODART = MG70_CODART_MG66" & vbCrLf & _
                    "LEFT OUTER JOIN MG53_FAMIGLIE  WITH (NOLOCK) ON MG53_DITTA_CG18 = MG70_DITTA_CG18" & vbCrLf & _
                    "    AND MG53_CODFAM = MG66_FAM_MG53" & vbCrLf & _
                    "LEFT OUTER JOIN MG54_SOTTOFAM  WITH (NOLOCK) ON MG54_DITTA_CG18 = MG70_DITTA_CG18" & vbCrLf & _
                    "    AND MG54_CODFAM_MG53 = MG66_FAM_MG53" & vbCrLf & _
                    "    AND MG54_CODSFAM = MG66_SFAM_MG54" & vbCrLf & _
                    "LEFT OUTER JOIN MG55_GRUPPI  WITH (NOLOCK) ON MG55_DITTA_CG18 = MG70_DITTA_CG18" & vbCrLf & _
                    "    AND MG55_CODFAM_MG53 = MG66_FAM_MG53" & vbCrLf & _
                    "    AND MG55_CODSFAM_MG54 = MG66_SFAM_MG54" & vbCrLf & _
                    "    AND MG55_CODGRUPPO = MG66_GRUPPO_MG55" & vbCrLf & _
                    "LEFT OUTER JOIN MG56_SOTTOGRUPPI  WITH (NOLOCK) ON MG56_DITTA_CG18 = MG70_DITTA_CG18" & vbCrLf
str_SQL = str_SQL & "    AND MG56_CODFAM_MG53 = MG66_FAM_MG53" & vbCrLf & _
                    "    AND MG56_CODSFAM_MG54 = MG66_SFAM_MG54" & vbCrLf & _
                    "    AND MG56_CODGRUPPO_MG55 = MG66_GRUPPO_MG55" & vbCrLf & _
                    "    AND MG56_CODSGRUPPO = MG66_SGRUPPO_MG56" & vbCrLf & _
                    "WHERE MG70_DITTA_CG18 = " & Gstr_DittaCorrente & vbCrLf & _
                    "    AND MG70_TIPOPROG = 1"
        Erase Pst_Colonne
        Pst_Colonne(0, 0) = "Codice Articolo"
        Pst_Colonne(0, 1) = ""
        Pst_Colonne(1, 0) = "Alias"
        Pst_Colonne(1, 1) = ""
        Pst_Colonne(2, 0) = "UM"
        Pst_Colonne(2, 1) = ""
        Pst_Colonne(3, 0) = "Fam."
        Pst_Colonne(3, 1) = ""
        Pst_Colonne(4, 0) = "Descrizione famiglia"
        Pst_Colonne(4, 1) = ""
        Pst_Colonne(5, 0) = "S/fam"
        Pst_Colonne(5, 1) = ""
        Pst_Colonne(6, 0) = "Descrizione sottofamiglia"
        Pst_Colonne(6, 1) = ""
        Pst_Colonne(7, 0) = "Gr."
        Pst_Colonne(7, 1) = ""
        Pst_Colonne(8, 0) = "Descrizione gruppo"
        Pst_Colonne(8, 1) = ""
        Pst_Colonne(9, 0) = "S/gr"
        Pst_Colonne(9, 1) = ""
        Pst_Colonne(10, 0) = "Descrizione sottogruppo"
        Pst_Colonne(10, 1) = ""
        
        Arr_Fields = Pst_Colonne
        Str_Caption = "Elenco articoli movimentati"
        Str_Connect = Gstr_Connect
        TXT_CODART.IDLookup = "Lkp_ArticoliMovimentati"
    End If
    
    
    Exit Sub
    
Err:
  Cancel = True
  Set Gcls_Log.vbError = Err
  Set Gcls_Log.ADOError = Gcon_Connect.Errors
  If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.TXT_CODART_StartLookup") = 1 Then
     Unload Me
  Else
     Resume Next
  End If
End Sub



Private Sub TXT_CODARTFOR_CloseLookup(Arr_Fields As Variant)
  TXT_CODART.Text = Arr_Fields(3, 1)
  
End Sub

Private Sub TXT_CODARTFOR_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
  On Error Resume Next
  
  Cancel = False
  
  'Enzo 200703 - Articolo fornitore
  str_SQL = "SELECT MG73_ARTCLIFOR, MG87_DESCART, MG73_CLIFOR_CG44, MG73_CODART_MG66 "
'  str_SQL = str_SQL & " FROM MG73_ARTCLIFOR"
  str_SQL = str_SQL & " FROM         MG87_ARTDESC RIGHT OUTER JOIN"
  str_SQL = str_SQL & "                       MG73_ARTCLIFOR ON MG87_ARTDESC.MG87_DITTA_CG18 = MG73_ARTCLIFOR.MG73_DITTA_CG18 AND"
  str_SQL = str_SQL & "                       MG87_ARTDESC.MG87_CODART_MG66 = MG73_ARTCLIFOR.MG73_CODART_MG66 AND "
  str_SQL = str_SQL & "                       MG87_ARTDESC.MG87_OPZIONE_MG5E = MG73_ARTCLIFOR.MG73_OPZIONE_MG5E"
  str_SQL = str_SQL & " WHERE MG73_DITTA_CG18 = " & Gstr_DittaCorrente
  
  If RTrimN(TXT_CODART.Text) <> "" And TXT_CODART.IsValid Then
    str_SQL = str_SQL & " AND MG73_CODART_MG66 = '" & TXT_CODART.Text & "'"
  End If
  
  str_SQL = str_SQL & " ORDER BY MG73_ARTCLIFOR ASC, MG73_FLGFORPREF DESC "
        
  ReDim Arr_Fields(0 To 3, 0 To 1)
  Arr_Fields(0, 0) = "Articolo Fornitore"
  Arr_Fields(0, 1) = ""
  Arr_Fields(1, 0) = "Descrizione"
  Arr_Fields(1, 1) = ""
  Arr_Fields(2, 0) = "Fornitore"
  Arr_Fields(2, 1) = ""
  Arr_Fields(3, 0) = "Articolo Interno"
  Arr_Fields(3, 1) = ""
  
  Str_Caption = "Articoli fornitori"
  Str_Connect = Gstr_Connect
  TXT_CODARTFOR.IDLookup = "lkp_ArtFor"
  
  Err.Clear
        
End Sub

'Private Sub TXT_CODART_AfterItem(Cancel As Boolean)
''    On Error GoTo Err
''
''    If Not TXT_OPZIONE.Enabled And RTrimN(TXT_CODART.Text) <> "" And TXT_CODART.IsValid Then
''        FME_CCS_SKPROD.ExecuteQuery
''    End If
''
''    Exit Sub
''
''Err:
''  Cancel = True
''  Set Gcls_Log.vbError = Err
''  Set Gcls_Log.ADOError = Gcon_Connect.Errors
''  If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.TXT_CODART_AfterItem") = 1 Then
''     Unload Me
''  Else
''     Resume Next
''  End If
'End Sub

'Private Sub TXT_OPZIONE_AfterItem(Cancel As Boolean)
''    On Error GoTo Err
''
''    If FME_CCS_SKPROD.Status = tsInsert Then
''        Call FME_CCS_SKPROD.ExecuteQuery
''        TXT_OPZIONE.Enabled = False
''    End If
''
''    Exit Sub
''
''Err:
''  Cancel = True
''  Set Gcls_Log.vbError = Err
''  Set Gcls_Log.ADOError = Gcon_Connect.Errors
''  If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.TXT_OPZIONE_AfterItem") = 1 Then
''     Unload Me
''  Else
''     Resume Next
''  End If
'End Sub

Private Sub TXT_OPZIONE_StartDecode(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Connect As String)
    On Error GoTo Err
    
    Dim Pst_Colonne(0 To 0, 0 To 1)     As Variant
    Erase Pst_Colonne
    Set Pst_Colonne(0, 0) = TXT_DESCART
    Pst_Colonne(0, 1) = "MG87_DESCART"
    Arr_Fields = Pst_Colonne
    Str_Connect = Gstr_Connect
    
    Exit Sub
Err:
  Cancel = True
  Set Gcls_Log.vbError = Err
  Set Gcls_Log.ADOError = Gcon_Connect.Errors
  If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.TXT_OPZIONE_StartDecode") = 1 Then
     Unload Me
  Else
     Resume Next
  End If
End Sub

'Private Sub CMB_TIPOQTA_ChangeRow(Cancel As Boolean, row As Integer)
''    On Error GoTo Err
''
''    Cancel = False
''
''    Call Psub_Elabora(RTrimN(TXT_CODART.Text), RTrimN(TXT_OPZIONE.Text))
''
''    Exit Sub
''
''Err:
''  Cancel = True
''  Set Gcls_Log.vbError = Err
''  Set Gcls_Log.ADOError = Gcon_Connect.Errors
''  If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.CMB_TIPOQTA_ChangeRow") = 1 Then
''     Unload Me
''  Else
''     Resume Next
''  End If
'End Sub

Public Sub Psub_Reinizializza()
    On Error GoTo Err:
    
    TXT_CODART.Text = ""
    TXT_DESCART.Text = ""
    TXT_DESCART.Default = ""
    
    'Enzo 200703 - Anagrafica estesa
    TXT_DESCARTEST.Text = ""
    TXT_DESCARTEST.Default = ""
    
    'Enzo 200703 - Carichi e scarichi
    TXT_TOTCARICHI.Text = ""
    TXT_TOTSCARICHI.Text = ""
    TXT_DATAULCA.Text = ""
    TXT_DATAULSCA.Text = ""
    
    TXT_INESAUR.Text = ""
    
    TXT_OPZIONE.Text = ""
    CMB_TIPOQTA.Text = 0
    CMB_TIPOART.Text = 4
    TXT_FAM.Text = ""
    TXT_SFAM.Text = ""
    TXT_GRUP.Text = ""
    TXT_SGRUP.Text = ""
    TXT_DESCFAM.Text = ""
    TXT_UM1.Text = ""
    TXT_UM2.Text = ""
    TXT_TIPOPROD.Text = ""
'    CHK_MOV.Text = 1
    
    Set GRID_GIACENZE.DataSource = Nothing
    GRID_GIACENZE.ReBind
    
    'Enzo 200703 - Ultimo prezzo acquisto e vendita
    Set GRID_LISVEN.DataSource = Nothing
    GRID_LISVEN.ReBind
    
    Set GRID_LISACQ.DataSource = Nothing
    GRID_LISACQ.ReBind
    
    Set GRID_LISACQ_TOT.DataSource = Nothing
    GRID_LISACQ_TOT.ReBind
    
    Set GRID_ARTALT.DataSource = Nothing
    GRID_ARTALT.ReBind
    
    Set GRID_ARTSOST.DataSource = Nothing
    GRID_ARTSOST.ReBind
    
    Exit Sub
    
Err:
  Set Gcls_Log.vbError = Err
  Set Gcls_Log.ADOError = Gcon_Connect.Errors
  If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.Psub_Reinizializza") = 1 Then
     Unload Me
  Else
     Resume Next
  End If
End Sub

Private Function ExecuteFormEvent(ByVal Mode As Variant)
Dim ClsScript                       As FWUO_TMSDEVELOP.CLSFW_PRSVBSCRIPT

    Select Case Mode
        Case "tsOpen"
            ActiveInterface.ActiveNavigator.InitializeScript
            Set ClsScript = ActiveInterface.ActiveNavigator.ClsScript
            If Not ClsScript Is Nothing Then
                ClsScript.ExecuteObjectEvent Me.Name, FWUO_TMSDEVELOP.tsForm, FWUO_TMSDEVELOP.tsCloseForm, Me.Name
            End If
        Case "tsClose"
            Set ClsScript = ActiveInterface.ActiveNavigator.ClsScript
            If Not ClsScript Is Nothing Then
                ClsScript.ExecuteObjectEvent Me.Name, FWUO_TMSDEVELOP.tsForm, FWUO_TMSDEVELOP.tsCloseForm, Me.Name
            End If
    End Select

End Function

Private Function SyncNavigator()
    On Error Resume Next
   
    Set ActiveInterface.ClsGlobal.ActiveInterface = ActiveInterface
    If Not ActiveInterface.ActiveFrame Is Nothing Then
        ActiveInterface.ActiveNavigator.SetStatus (ActiveInterface.ActiveFrame.Status)
    End If
    ActiveInterface.ActiveNavigator.Refresh

End Function

Public Function RTrimN(ByVal fvar_value As Variant) As Variant
    If IsNull(fvar_value) Or IsEmpty(fvar_value) Or (fvar_value = "") Then
        RTrimN = ""
    Else
        RTrimN = RTrim(fvar_value)
    End If
    
End Function

Private Function NVL(ByVal Valore As Variant)
    On Error Resume Next
   
    If RTrimN(Valore) = "" Then
       NVL = 0
    Else
       NVL = Valore
    End If
   
    Err.Clear
    
End Function

Private Function ProgressiviProgetto() As Boolean

    On Error GoTo Err

    Dim Sql_prog                            As String
    Dim Rst_prog                            As ADODB.Recordset
    
    Sql_prog = "SELECT * FROM MG4F_PARAMMOVLOTTI WITH (NOLOCK) " & _
                " WHERE MG4F_DITTA_CG18 = " & Gstr_DittaCorrente & _
                " AND MG4F_TIPOMOV like '%' + 'PRO' + '%'"

    Set Rst_prog = Gcon_Connect.Execute(Sql_prog)

    If Not Rst_prog Is Nothing Then
        If Rst_prog.RecordCount > 0 Then
            ProgressiviProgetto = True
            NumProg = Rst_prog("MG4F_PROG").Value
        Else
            ProgressiviProgetto = False
            NumProg = 0
        End If
    End If

    Exit Function

Err:
  Screen.MousePointer = vbDefault
  ActiveInterface.StatusBar.Panels(2) = "Pronto"
  Set Gcls_Log.vbError = Err
  Set Gcls_Log.ADOError = Gcon_Connect.Errors
  If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.ProgressiviProgetto") = 1 Then
     Unload Me
  Else
     Resume Next
  End If
  
End Function

Private Sub Formatta_colonne()

Dim Pint_count          As Integer

    Set cls_datagrid = New CLSFW_DataGrid
    Set cls_datagrid.ActiveInterface = ActiveInterface
    Set cls_datagrid.DataGrid = GRID_GIACENZE
        
    For Pint_count = 0 To GRID_GIACENZE.Columns.Count - 1
        Select Case GRID_GIACENZE.Columns(Pint_count).DataField
            Case "COD_PROGETTO"
                'Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, Generico, tsGenerico)
            Case "QGIACINI"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QGIACATT"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QGIACEFF"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QGIACFIS"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QDISPONIB"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QIMPCLI"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QIMPPROD"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QIMPCLAVFOR"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QPREIMPCLI"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QBLOCSPED"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QDACONTR"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QORDFOR"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QORDPROD"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QPREIMPFOR"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QDAVAL"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QENTCVIS"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QENTCRIP"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QENCDEP"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QENCNOLO"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QUSCCVIS"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QUSCCRIP"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QUSCDEP"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QUSCNOLO"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QCARACQ"
                  Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QCARESORCLI"
                  Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QCARPROD"
                  Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QCARCLAVCLI"
                  Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QCARCLAVFOR"
                  Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QCAROMAG"
                  Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QCARGENER"
                  Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QCARTRASF"
                  Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QCARSOST"
                  Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QCARLIB1"
                  Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QCARLIB2"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QSCAVEN"
                  Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QSCASCART"
                  Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QSCAOMAGQ"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QSCACLAVCLI"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QSCACLAVFOR"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QSCAPROD"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QSCARESOFOR"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QSCAGENER"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QSCATRASF"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QSCASOST"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QSCALIB1"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
            Case "QSCALIB2"
                Call cls_datagrid.AddColumn(GRID_GIACENZE.Columns(Pint_count).DataField, QuantitaMovimentata, tsQuantita1, "EURO")
        End Select
    Next
    
End Sub

Public Function Decimali() As Integer

On Error GoTo Err

    Dim Sql_Dec                                         As String
    Dim Rst_Dec                                         As ADODB.Recordset
    
    Sql_Dec = "SELECT MG66_INDDECQTA FROM MG66_ANAGRART WITH (NOLOCK) " & _
                " WHERE MG66_DITTA_CG18 = " & Gstr_DittaCorrente & _
                " AND MG66_CODART = '" & RTrimN(TXT_CODART.Text) & "'"

    Set Rst_Dec = Gcon_Connect.Execute(Sql_Dec, , adCmdText)

'       RICCI ROBERTO 08/05/2006
'
'   Commento e lascio prendere sempre i decimali anche se = zero
'
'    If Not Rst_Dec.EOF Or Not Rst_Dec.BOF Then
'           Decimali = CDec(Rst_Dec("MG66_INDDECQTA").Value)
'        Else
'           Decimali = 3
'        End If
'    End If

    If Not Rst_Dec.EOF Or Not Rst_Dec.BOF Then
        pvarDecimali = CDec(Rst_Dec("MG66_INDDECQTA").Value)
    End If
    
    Exit Function

Err:
  Set Gcls_Log.vbError = Err
  Set Gcls_Log.ADOError = Gcon_Connect.Errors
  If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.Decimali") = 1 Then
     Unload Me
  Else
     Resume Next
  End If

End Function

Private Sub ReinizializzaVirtualFrame()

On Error GoTo Err

Dim str_string                      As String

    str_string = " SELECT" & _
                "    MG66_CODART" & _
                " FROM" & _
                "    MG66_ANAGRART WITH (NOLOCK) " & _
                " WHERE" & _
                "    MG66_DITTA_CG18 = " & Gstr_DittaCorrente
    Grst_SitGiacenze.Close
    Grst_SitGiacenze.Open str_string & " AND 1=0 ORDER BY MG66_CODART"
    FME_CCS_SKPROD.ReOpen str_string & " ORDER BY MG66_CODART"

    Exit Sub

Err:
  Set Gcls_Log.vbError = Err
  Set Gcls_Log.ADOError = Gcon_Connect.Errors
  If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.ReinizializzaVirtualFrame") = 1 Then
     Unload Me
  Else
     Resume Next
  End If

End Sub
Private Sub InvocaSkPrezzi(ByVal IndVenAcq As Integer)
'Dim SkPrezzi_Interface As Cinterface
'
'    On Error GoTo Err
'
'    Set Cls_ConnectMagazzino.ActiveInterface = ActiveInterface
'    Cls_ConnectMagazzino.Left = 50
'    Cls_ConnectMagazzino.Top = 1000
'    Call Cls_ConnectMagazzino.SchedaPrezziArticoli(RTrimN(TXT_CODART.Text), RTrimN(TXT_OPZIONE.Text), IndVenAcq)
'    ActiveInterface.IsActive = True
'    Set Cls_ConnectMagazzino.ActiveInterface = Nothing
'    Set ActiveInterface.ClsGlobal.ActiveInterface = ActiveInterface
'    Set ActiveInterface.ActiveFrame = FME_CCS_SKPROD
'    SyncNavigator
'    ActiveInterface.ActiveNavigator.InitializeScript
    
'    Me.Hide
'
'    Set Pcls_SkPrezzi = New CLSMG_SCHEDAPRZART
'    Set SkPrezzi_Interface = Pcls_SkPrezzi
'
'    Set Pcls_SkPrezzi.Chiamante = ActiveInterface
'
'    ActiveInterface.IsActive = False
'    Set ActiveInterface.ClsGlobal.ActiveInterface = SkPrezzi_Interface
'    ActiveInterface.ClsGlobal.ActiveInterface.IsActive = True
'
'    Set ActiveInterface.ClsGlobal.CallInterface = SkPrezzi_Interface
'    SkPrezzi_Interface.IsCalled = True
'
'    Pcls_SkPrezzi.CodiceArticolo = RTrimN(TXT_CODART.Text)
'    Pcls_SkPrezzi.Opzione = RTrimN(TXT_OPZIONE.Text)
'    Pcls_SkPrezzi.AcquistoVendita = IndVenAcq
'    ActiveInterface.ClsGlobal.ExecDll False, "MGUO_SCHEDAPRZART.CLSMG_SCHEDAPRZART", False, tsInsert, Normale, 0, 0
'
'    Set ActiveInterface.ClsGlobal.ActiveInterface = Nothing
'    Set SkPrezzi_Interface = Nothing
'    Set Pcls_SkPrezzi = Nothing
'    ActiveInterface.IsActive = True
'    Set ActiveInterface.ClsGlobal.ActiveInterface = ActiveInterface
'    Set ActiveInterface.ClsGlobal.Gcls_VoceMenu = ActiveInterface.ClsVoceMenu
'
'    Set ActiveInterface.ActiveFrame = FME_CCS_SKPROD
'    SyncNavigator
'    ActiveInterface.ActiveNavigator.InitializeScript
    
    Exit Sub

Err:
    Set Gcls_Log.vbError = Err
    Set Gcls_Log.ADOError = Gcon_Connect.Errors
    If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.InvocaSkPrezzi") = 1 Then
       Unload Me
    Else
       Resume Next
    End If
End Sub

Private Sub InvocaVerificaGiacenzeCLavoro()
Dim Partitario_Interface As Cinterface
    
    On Error GoTo Err

#If Not GAMMA_SPRINT Then
    Set Pcls_Connect_Produzione.ActiveInterface = ActiveInterface
    
    Pcls_Connect_Produzione.CodiceArticolo = RTrimN(TXT_CODART.Text)
    Pcls_Connect_Produzione.CodiceVarianteArticolo = RTrimN(TXT_OPZIONE.Text)
    'Setto il combo Modalità ordinemento dati per Articolo = 1
'    Pcls_Connect_Produzione.ModalitaOrdinamentoDati = 1
'    Call Pcls_Connect_Produzione.CallGiacenzeContoLavoro
    ActiveInterface.IsActive = True
    Set Pcls_Connect_Produzione.ActiveInterface = Nothing
    Set ActiveInterface.ClsGlobal.ActiveInterface = ActiveInterface
    Set ActiveInterface.ActiveFrame = FME_CCS_SKPROD
    SyncNavigator
    ActiveInterface.ActiveNavigator.InitializeScript

#End If

    Exit Sub

Err:
    Set Gcls_Log.vbError = Err
    Set Gcls_Log.ADOError = Gcon_Connect.Errors
    If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.InvocaVerificaGiacenzeCLavoro") = 1 Then
       Unload Me
    Else
       Resume Next
    End If
End Sub

Public Sub ImpostaVirtualFrame(ByVal Operazione As Integer)

On Error GoTo Err

'
'Operazione:
'0-> Record successivo
'1-> Record precedente
'2-> Primo record
'3-> Ultimo record

Dim str_string                      As String
Dim Rst_test                        As ADODB.Recordset

'    Set Rst_test = New ADODB.Recordset
    
    Select Case Operazione
        Case 0
            'Prendo il record successivo
            str_string = " SELECT TOP 1 " & _
                        "    MG66_CODART" & _
                        " FROM" & _
                        "    MG66_ANAGRART WITH (NOLOCK) " & _
                        " WHERE" & _
                        "    MG66_DITTA_CG18 = " & Gstr_DittaCorrente & _
                        "   AND MG66_CODART > '" & RTrimN(TXT_CODART.Text) & "'" & _
                        "   ORDER BY MG66_CODART"
            Set Rst_test = Gcon_Connect.Execute(str_string)
            If Rst_test.RecordCount = 0 Then
                Exit Sub
            End If
        Case 1
            'Prendo il record precedente
            str_string = " SELECT TOP 1 " & _
                        "    MG66_CODART" & _
                        " FROM" & _
                        "    MG66_ANAGRART WITH (NOLOCK) " & _
                        " WHERE" & _
                        "    MG66_DITTA_CG18 = " & Gstr_DittaCorrente & _
                        "   AND MG66_CODART < '" & RTrimN(TXT_CODART.Text) & "'" & _
                        "   ORDER BY MG66_CODART DESC"
            Set Rst_test = Gcon_Connect.Execute(str_string)
            If Rst_test.RecordCount = 0 Then
                Exit Sub
            End If
        Case 2
            'Prendo il primo record
            str_string = " SELECT TOP 1 " & _
                        "    MG66_CODART" & _
                        " FROM" & _
                        "    MG66_ANAGRART WITH (NOLOCK) " & _
                        " WHERE" & _
                        "    MG66_DITTA_CG18 = " & Gstr_DittaCorrente & _
                        "   ORDER BY MG66_CODART"
            Set Rst_test = Gcon_Connect.Execute(str_string)
            If Rst_test.RecordCount = 0 Then
                Exit Sub
            End If
        Case 3
            'Prendo il l'ultimo record
            str_string = " SELECT TOP 1 " & _
                        "    MG66_CODART" & _
                        " FROM" & _
                        "    MG66_ANAGRART WITH (NOLOCK) " & _
                        " WHERE" & _
                        "    MG66_DITTA_CG18 = " & Gstr_DittaCorrente & _
                        "   ORDER BY MG66_CODART DESC"
            Set Rst_test = Gcon_Connect.Execute(str_string)
            If Rst_test.RecordCount = 0 Then
                Exit Sub
            End If
    End Select
            
    Grst_SitGiacenze.Close
    Grst_SitGiacenze.Open str_string
    FME_CCS_SKPROD.ReOpen str_string
    Call RiempioDati(RTrimN(Grst_SitGiacenze.Fields("MG66_CODART").Value), "")
    Exit Sub

Err:
  Set Gcls_Log.vbError = Err
  Set Gcls_Log.ADOError = Gcon_Connect.Errors
  If Gcls_Log.ViewRunTimeErr("0_0_0_0", "MGUO_SITGIACENZE", "FRMMG_SITGIACENZE.ImpostaVirtualFrame") = 1 Then
     Unload Me
  Else
     Resume Next
  End If

End Sub
Private Function VisualizzaErrore(ByVal SubOrFunctionName As String) As VbMsgBoxResult
    '
    ' setto l'oggetto errore di VB
    '
    Set Gcls_Log.vbError = Err
    '
    ' setto l'eventuale errore ADO
    '
    If Not (Gcon_Connect Is Nothing) Then
        Set Gcls_Log.ADOError = Gcon_Connect.Errors
    End If
    '
    ' invoco il metodo di visualizzazione dell'errore
    '
    VisualizzaErrore = Gcls_Log.ShowError(App.Title, Me.Caption, SubOrFunctionName)
End Function

Private Function CDecN(ByVal fvar_value As Variant) As Variant
    If IsNull(fvar_value) Or IsEmpty(fvar_value) Or (fvar_value = "") Then
        CDecN = 0
    Else
        CDecN = CDec(fvar_value)
    End If
End Function
