VERSION 5.00
Object = "{0EF4EAA6-2617-11D2-A1C0-0060082875F9}#4.14#0"; "TMS_COMBOBOX.ocx"
Object = "{0EF4EA3A-2617-11D2-A1C0-0060082875F9}#8.12#0"; "TMS_EDIT.ocx"
Object = "{5032AB27-52C8-11D2-A1C0-0060082875F9}#4.23#0"; "TMS_EDITM.ocx"
Object = "{0EF4E9DB-2617-11D2-A1C0-0060082875F9}#10.13#0"; "TMS_EDITNUM.ocx"
Object = "{F53BE214-7AC6-11D0-9B0E-006097A80EFD}#6.12#0"; "TMS_LABEL.ocx"
Object = "{0EF4E915-2617-11D2-A1C0-0060082875F9}#7.22#0"; "TMS_RICHTEXTBOX.ocx"
Object = "{31930FDA-530C-11D2-A1C0-0060082875F9}#2.35#0"; "TMS_ARTICOLO.ocx"
Object = "{9AE03505-25F7-11D2-A1C0-0060082875F9}#7.3#0"; "TMS_FRAME.ocx"
Object = "{CBAF6F53-3C3D-11D4-AA70-000629C16DEA}#2.4#0"; "MDIActiveXS.ocx"
Object = "{B473387D-A75F-4A83-9879-4A8FE48EE80F}#1.8#0"; "TMS_TBARMENU.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{840F600B-FE39-42F4-AE87-798701D999E2}#1.23#0"; "TMS_RESIZEFORM.ocx"
Object = "{EF28CC5E-FCE3-448A-AB46-AEA7C5A209AA}#1.4#0"; "TMS_SSTAB.ocx"
Object = "{53EEE555-1204-4E18-B5DB-A659E06A9EEB}#1.3#0"; "TMS_FLATBUTTON.ocx"
Object = "{C217CF55-DAD6-4868-A146-622ECD75BC60}#1.60#0"; "TMS_QGRID.ocx"
Object = "{5CC9FF70-1720-11D2-A1C0-0060082875F9}#3.6#0"; "TMS_GRIDNAV.ocx"
Begin VB.Form FRMIT_CONFDIST 
   Caption         =   "Pannello Analisi Mansco"
   ClientHeight    =   13290
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   21015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMIT_CONFDIST.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   13290
   ScaleWidth      =   21015
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   6660
      TabIndex        =   98
      Top             =   1890
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin PRJ_SSTAB.TMS_SSTAB TAB_DISTINTA 
      Height          =   11190
      Left            =   15
      TabIndex        =   61
      Top             =   2490
      Width           =   21555
      _ExtentX        =   38021
      _ExtentY        =   19738
      TabCount        =   4
      TabCaption(0)   =   "SEMILAVORATI"
      TabContCtrlCnt(0)=   1
      Tab(0)ContCtrlCap(1)=   "TMS_ANAGRAFICI"
      TabCaption(1)   =   "IMBALLI"
      TabContCtrlCnt(1)=   1
      Tab(1)ContCtrlCap(1)=   "TMS_FRAME5"
      TabCaption(2)   =   "FASI"
      TabContCtrlCnt(2)=   2
      Tab(2)ContCtrlCap(1)=   "TMS_FRAME9"
      Tab(2)ContCtrlCap(2)=   "TMS_FRAME8"
      TabCaption(3)   =   "GENERAZIONE DISTINTA"
      TabContCtrlCnt(3)=   1
      Tab(3)ContCtrlCap(1)=   "TXT_LOG"
      ActiveTab       =   2
      Begin PRJFW_RICHTEXTBOX.TmsRichTextBox TXT_LOG 
         Height          =   10200
         Left            =   -75135
         TabIndex        =   97
         Top             =   405
         Width           =   21615
         _ExtentX        =   38126
         _ExtentY        =   17992
         MaxChar         =   10000
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         NumRighe        =   34
         IsMultiLine     =   -1  'True
         MaxWidth        =   178
         IsInLingua      =   0   'False
         LinguaEntitaDes =   0
         LinguaIDProvenienzaExt=   ""
      End
      Begin PRJFW_FRAME.TMS_FRAME TMS_FRAME9 
         Height          =   10455
         Left            =   9075
         Top             =   390
         Width           =   8970
         _ExtentX        =   15822
         _ExtentY        =   18441
         Begin TMS_QGRID.TMS_QGRIDWRAPPER QGRID_FASISL 
            Height          =   10050
            Left            =   45
            TabIndex        =   84
            Top             =   375
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   17727
            ColorTheme      =   0
            BeginProperty PreviewFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PreviewFontColor=   0
         End
         Begin PRJFW_GRIDNAV.TMS_GRIDNAV GRIDNAVFASISL 
            Height          =   360
            Left            =   45
            TabIndex        =   89
            Top             =   45
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   635
            Primo           =   0   'False
            Precedente      =   0   'False
            Successivo      =   0   'False
            Ultimo          =   0   'False
         End
         Begin PRJFW_EDITM.TXT_EDITM TXT_MACCHINA_FSL 
            Height          =   300
            Left            =   6975
            TabIndex        =   101
            Top             =   30
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   529
            IsLookup        =   -1  'True
            DisplayFormat   =   "Maiuscolo"
            MaxChar         =   25
            Numerico        =   0   'False
            Carattere       =   0   'False
            DBField         =   "IT01_MACCHINA_PD08"
            NumRighe        =   0
            IsInLingua      =   0   'False
            LinguaEntitaDes =   0
            LinguaIDProvenienzaExt=   ""
         End
         Begin PRJFW_EDITNUM.TxtEditNum TXT_SEQSL 
            Height          =   300
            Left            =   5700
            TabIndex        =   94
            Top             =   735
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   529
            DBField         =   "IT01_SEQ"
         End
         Begin PRJFW_EDIT.TxtEdit TXT_REPFASESL 
            Height          =   300
            Left            =   1215
            TabIndex        =   92
            Top             =   2340
            Width           =   1680
            _ExtentX        =   2566
            _ExtentY        =   529
            Numerico        =   0   'False
            Carattere       =   0   'False
            DBField         =   "IT01_CODREP_PD07"
         End
         Begin PRJFW_EDIT.TxtEdit TXT_DESCRFASESL 
            Height          =   300
            Left            =   105
            TabIndex        =   91
            Top             =   1890
            Width           =   1680
            _ExtentX        =   2566
            _ExtentY        =   529
            Numerico        =   0   'False
            Carattere       =   0   'False
            DBField         =   "IT01_DESCR"
         End
         Begin PRJFW_EDIT.TxtEdit TXT_FASESL 
            Height          =   300
            Left            =   210
            TabIndex        =   90
            Top             =   1140
            Width           =   1680
            _ExtentX        =   2566
            _ExtentY        =   529
            Numerico        =   0   'False
            Carattere       =   0   'False
            DBField         =   "IT01_CODFASE"
         End
      End
      Begin PRJFW_FRAME.TMS_FRAME TMS_FRAME8 
         Height          =   10455
         Left            =   90
         Top             =   390
         Width           =   8970
         _ExtentX        =   15822
         _ExtentY        =   18441
         Begin PRJFW_EDITM.TXT_EDITM TXT_MACCHINA_FPF 
            Height          =   300
            Left            =   6945
            TabIndex        =   100
            Top             =   30
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   529
            IsLookup        =   -1  'True
            DisplayFormat   =   "Maiuscolo"
            MaxChar         =   25
            Numerico        =   0   'False
            Carattere       =   0   'False
            DBField         =   "IT01_MACCHINA_PD08"
            NumRighe        =   0
            IsInLingua      =   0   'False
            LinguaEntitaDes =   0
            LinguaIDProvenienzaExt=   ""
         End
         Begin TMS_QGRID.TMS_QGRIDWRAPPER QGRID_FASIPF 
            Height          =   10080
            Left            =   15
            TabIndex        =   83
            Top             =   345
            Width           =   8880
            _ExtentX        =   15663
            _ExtentY        =   17780
            ColorTheme      =   0
            BeginProperty PreviewFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PreviewFontColor=   0
         End
         Begin PRJFW_GRIDNAV.TMS_GRIDNAV GRIDNAVFASIPF 
            Height          =   360
            Left            =   0
            TabIndex        =   85
            Top             =   0
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   635
            Primo           =   0   'False
            Precedente      =   0   'False
            Successivo      =   0   'False
            Ultimo          =   0   'False
         End
         Begin PRJFW_EDITNUM.TxtEditNum TXT_SEQPF 
            Height          =   300
            Left            =   6435
            TabIndex        =   93
            Top             =   795
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   529
            DBField         =   "IT01_SEQ"
         End
         Begin PRJFW_EDIT.TxtEdit TXT_REPFASEPF 
            Height          =   300
            Left            =   6015
            TabIndex        =   88
            Top             =   1785
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            Numerico        =   0   'False
            Carattere       =   0   'False
            DBField         =   "IT01_CODREP_PD07"
         End
         Begin PRJFW_EDIT.TxtEdit TXT_DESCRFASEPF 
            Height          =   300
            Left            =   5820
            TabIndex        =   87
            Top             =   1215
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            Numerico        =   0   'False
            Carattere       =   0   'False
            DBField         =   "IT01_DESCR"
         End
         Begin PRJFW_EDIT.TxtEdit TXT_FASEPF 
            Height          =   300
            Left            =   5805
            TabIndex        =   86
            Top             =   390
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            Numerico        =   0   'False
            Carattere       =   0   'False
            DBField         =   "IT01_CODFASE"
         End
      End
      Begin PRJFW_FRAME.TMS_FRAME TMS_FRAME5 
         Height          =   10380
         Left            =   -75000
         Top             =   375
         Width           =   21480
         _ExtentX        =   37888
         _ExtentY        =   18309
         Begin PRJFW_FRAME.TMS_FRAME TMS_FRAME7 
            Height          =   10320
            Left            =   8445
            Top             =   30
            Width           =   8130
            _ExtentX        =   14340
            _ExtentY        =   18203
            Begin TMS_QGRID.TMS_QGRIDWRAPPER QGRID_IMBALLIPF 
               Height          =   10020
               Left            =   30
               TabIndex        =   74
               Top             =   255
               Width           =   8055
               _ExtentX        =   14208
               _ExtentY        =   17674
               ColorTheme      =   0
               BeginProperty PreviewFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PreviewFontColor=   0
            End
            Begin PRJFW_GRIDNAV.TMS_GRIDNAV GRIDNAVIMBALLIPF 
               Height          =   360
               Left            =   6450
               TabIndex        =   75
               Top             =   9915
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   635
               Primo           =   0   'False
               Precedente      =   0   'False
               Successivo      =   0   'False
               Ultimo          =   0   'False
            End
            Begin PRJFW_EDITNUM.TxtEditNum TXT_PesoAsciuttoImballiPF 
               Height          =   300
               Left            =   5700
               TabIndex        =   82
               Top             =   9285
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   529
               DBField         =   "PesoAsciutto"
            End
            Begin PRJFW_EDITNUM.TxtEditNum TXT_PesoUmidoImballiPF 
               Height          =   300
               Left            =   5730
               TabIndex        =   81
               Top             =   8700
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   529
               DBField         =   "PesoUmido"
            End
            Begin PRJFW_EDITNUM.TxtEditNum TXT_NumeroTubiImballiPF 
               Height          =   300
               Left            =   5670
               TabIndex        =   80
               Top             =   7785
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   529
               DBField         =   "NumeroTubi"
            End
            Begin PRJFW_EDITNUM.TxtEditNum TXT_QuantitaImballiPF 
               Height          =   300
               Left            =   5430
               TabIndex        =   79
               Top             =   6765
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   529
               DBField         =   "Quantita"
            End
            Begin PRJFW_EDITNUM.TxtEditNum TXT_PesoNImballiPF 
               Height          =   300
               Left            =   5520
               TabIndex        =   78
               Top             =   4875
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   529
               DBField         =   "PesoN"
            End
            Begin PRJFW_EDIT.TxtEdit TXT_DescrizioneArticoloImballiPF 
               Height          =   300
               Left            =   5505
               TabIndex        =   77
               Top             =   2865
               Width           =   2055
               _ExtentX        =   3625
               _ExtentY        =   529
               MaxChar         =   75
               Numerico        =   0   'False
               Carattere       =   0   'False
               DBField         =   "DescrizioneArticolo"
               MaxWidth        =   15
            End
            Begin PRJFW_EDIT.TxtEdit TXT_articoloimballiPF 
               Height          =   300
               Left            =   4650
               TabIndex        =   76
               Top             =   1575
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   529
               MaxChar         =   25
               Numerico        =   0   'False
               Carattere       =   0   'False
               DBField         =   "Articolo"
               MaxWidth        =   25
            End
            Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON TMS_FLATBUTTON3 
               Height          =   210
               Left            =   0
               TabIndex        =   73
               Top             =   30
               Width           =   8160
               _ExtentX        =   14393
               _ExtentY        =   370
               Caption         =   "I M B A L L I  P R O D O T T O  F I N I T O"
               ButtonBorderColor=   0
               ButtonForeColor =   16777215
               ButtonHilightBorderColor=   9473677
            End
         End
         Begin PRJFW_FRAME.TMS_FRAME TMS_FRAME6 
            Height          =   14490
            Left            =   30
            Top             =   45
            Width           =   8235
            _ExtentX        =   14526
            _ExtentY        =   25559
            Begin TMS_QGRID.TMS_QGRIDWRAPPER QGRID_IMBALLI 
               Height          =   10230
               Left            =   15
               TabIndex        =   71
               Top             =   225
               Width           =   8145
               _ExtentX        =   14367
               _ExtentY        =   18045
               ColorTheme      =   0
               BeginProperty PreviewFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PreviewFontColor=   0
            End
            Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON TMS_FLATBUTTON1 
               Height          =   210
               Left            =   0
               TabIndex        =   96
               Top             =   -15
               Width           =   8160
               _ExtentX        =   14393
               _ExtentY        =   370
               Caption         =   "I M B A L L I  P R O D O T T O S L"
               ButtonBorderColor=   0
               ButtonForeColor =   16777215
               ButtonHilightBorderColor=   9473677
            End
            Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON TMS_FLATBUTTON4 
               Height          =   210
               Left            =   -390
               TabIndex        =   72
               Top             =   -4470
               Width           =   8160
               _ExtentX        =   14393
               _ExtentY        =   370
               Caption         =   "I M B A L L I"
               ButtonBorderColor=   0
               ButtonForeColor =   16777215
               ButtonHilightBorderColor=   9473677
            End
         End
         Begin PRJFW_GRIDNAV.TMS_GRIDNAV GRIDNAVIMBALLI 
            Height          =   360
            Left            =   5940
            TabIndex        =   63
            Top             =   9390
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   635
            Primo           =   0   'False
            Precedente      =   0   'False
            Successivo      =   0   'False
            Ultimo          =   0   'False
         End
         Begin PRJFW_EDITNUM.TxtEditNum TXT_PesoNImballi 
            Height          =   300
            Left            =   630
            TabIndex        =   70
            Top             =   2340
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   529
            DBField         =   "PesoN"
         End
         Begin PRJFW_EDITNUM.TxtEditNum TXT_QuantitaImballi 
            Height          =   300
            Left            =   705
            TabIndex        =   69
            Top             =   2760
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   529
            DBField         =   "Quantita"
         End
         Begin PRJFW_EDITNUM.TxtEditNum TXT_NumeroTubiImballi 
            Height          =   300
            Left            =   810
            TabIndex        =   68
            Top             =   3210
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   529
            DBField         =   "NumeroTubi"
         End
         Begin PRJFW_EDITNUM.TxtEditNum TXT_PesoUmidoImballi 
            Height          =   300
            Left            =   945
            TabIndex        =   67
            Top             =   3795
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   529
            DBField         =   "PesoUmido"
         End
         Begin PRJFW_EDITNUM.TxtEditNum TXT_PesoAsciuttoImballi 
            Height          =   300
            Left            =   915
            TabIndex        =   66
            Top             =   4260
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   529
            DBField         =   "PesoAsciutto"
         End
         Begin PRJFW_EDIT.TxtEdit TXT_articoloimballi 
            Height          =   300
            Left            =   240
            TabIndex        =   65
            Top             =   870
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   529
            MaxChar         =   25
            Numerico        =   0   'False
            Carattere       =   0   'False
            DBField         =   "Articolo"
            MaxWidth        =   25
         End
         Begin PRJFW_EDIT.TxtEdit TXT_DescrizioneArticoloImballi 
            Height          =   300
            Left            =   390
            TabIndex        =   64
            Top             =   1395
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   529
            MaxChar         =   75
            Numerico        =   0   'False
            Carattere       =   0   'False
            DBField         =   "DescrizioneArticolo"
            MaxWidth        =   15
         End
      End
      Begin PRJFW_FRAME.TMS_FRAME TMS_ANAGRAFICI 
         Height          =   9825
         Left            =   -74970
         Top             =   375
         Width           =   21480
         _ExtentX        =   37888
         _ExtentY        =   17330
         Begin TMS_QGRID.TMS_QGRIDWRAPPER QGRID_ARTICOLI 
            Height          =   9750
            Left            =   45
            TabIndex        =   62
            Top             =   30
            Width           =   21360
            _ExtentX        =   37677
            _ExtentY        =   17198
            ColorTheme      =   0
            ScrollBar       =   0
            BeginProperty PreviewFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PreviewFontColor=   0
         End
      End
   End
   Begin PRJFW_FRAME.TMS_FRAME TMS_FRAME1 
      Height          =   1035
      Left            =   30
      Top             =   0
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   1826
      Begin PRJFW_ARTICOLO.TxtArticolo TXT_CODART 
         Height          =   300
         Left            =   870
         TabIndex        =   0
         ToolTipText     =   "Articolo"
         Top             =   60
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   529
         MaxChar         =   25
         Obbligatorio    =   -1  'True
         Numerico        =   0   'False
         Carattere       =   0   'False
         DBField         =   "MG66_CODART"
         IsQbe           =   -1  'True
         IsDecode        =   -1  'True
         Caption         =   "Articolo"
         Object.Tag             =   "Articolo"
         MaxWidth        =   15
         CanReturnRecordSet=   -1  'True
      End
      Begin MDIinActiveX.MDIActiveX MDIActiveX1 
         Left            =   7080
         Top             =   510
         _ExtentX        =   847
         _ExtentY        =   794
         BorderStyle     =   0
      End
      Begin PRJ_RESIZEFORM.TMS_RESIZEFORM TMS_RESIZEFORM1 
         Left            =   7650
         Top             =   510
         _ExtentX        =   847
         _ExtentY        =   847
      End
      Begin PRJFW_RICHTEXTBOX.TmsRichTextBox TXT_DESCART 
         Height          =   300
         Left            =   30
         TabIndex        =   1
         ToolTipText     =   "Descrizione articolo"
         Top             =   405
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
         IsInLingua      =   0   'False
         LinguaEntitaDes =   0
         LinguaIDProvenienzaExt=   0
      End
      Begin PRJFW_TBARMENU.TMS_TBARMENU CMD_NUOVO 
         Height          =   345
         Left            =   5220
         TabIndex        =   2
         Top             =   60
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         Caption         =   "&Nuovo"
         IsMenuPopup     =   0   'False
      End
      Begin PRJFW_TBARMENU.TMS_TBARMENU CMD_ELABORA 
         Height          =   345
         Left            =   4095
         TabIndex        =   3
         Top             =   75
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "&Elabora"
         IsMenuPopup     =   0   'False
      End
      Begin PRJFW_RICHTEXTBOX.TmsRichTextBox TXT_DESCARTEST 
         Height          =   300
         Left            =   30
         TabIndex        =   19
         ToolTipText     =   "Descrizione articolo"
         Top             =   705
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
         IsInLingua      =   0   'False
         LinguaEntitaDes =   0
         LinguaIDProvenienzaExt=   0
      End
      Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON TMS_FLATBUTTON2 
         Height          =   345
         Left            =   6375
         TabIndex        =   45
         Top             =   60
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   609
         ButtonBorderColor=   0
         ButtonForeColor =   16777215
         ButtonHilightBorderColor=   9473677
      End
      Begin PRJFW_TmsLabel.TMS_LABEL LBL_CODART 
         Height          =   300
         Left            =   240
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   90
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         Caption         =   "Articolo"
      End
   End
   Begin PRJFW_FRAME.TMS_FRAME TMS_FRAME3 
      Height          =   2430
      Left            =   8265
      Top             =   15
      Width           =   13350
      _ExtentX        =   23548
      _ExtentY        =   4286
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL15 
         Height          =   300
         Left            =   3135
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1890
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Caption         =   "Tot. Colla ( Kg.)"
      End
      Begin PRJFW_EDIT.TxtEdit TXT_TOTCOLLA 
         Height          =   300
         Left            =   4305
         TabIndex        =   58
         ToolTipText     =   "PZ"
         Top             =   1860
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
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL14 
         Height          =   300
         Left            =   9480
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1860
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   529
         Caption         =   "Peso ( kg / m ) Asciutto"
      End
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL13 
         Height          =   300
         Left            =   9465
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   1515
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   529
         Caption         =   "Peso ( kg / m ) Umido"
      End
      Begin PRJFW_EDIT.TxtEdit TXT_PESOA 
         Height          =   300
         Left            =   11160
         TabIndex        =   55
         ToolTipText     =   "PZ"
         Top             =   1830
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
      Begin PRJFW_EDIT.TxtEdit TXT_PESOU 
         Height          =   300
         Left            =   11175
         TabIndex        =   54
         ToolTipText     =   "PZ"
         Top             =   1470
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
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL11 
         Height          =   300
         Left            =   6105
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1875
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         Caption         =   "Peso Tubo Kg Asciutto"
      End
      Begin PRJFW_EDIT.TxtEdit TXT_PESOTUBOA 
         Height          =   300
         Left            =   7575
         TabIndex        =   52
         ToolTipText     =   "PZ"
         Top             =   1845
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
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL8 
         Height          =   300
         Left            =   6105
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1545
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         Caption         =   "Peso Tubo Kg Umido"
      End
      Begin PRJFW_EDIT.TxtEdit TXT_PESOTUBOU 
         Height          =   300
         Left            =   7575
         TabIndex        =   50
         ToolTipText     =   "PZ"
         Top             =   1500
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
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL5 
         Height          =   300
         Left            =   6105
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   855
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         Caption         =   "Tot. Carte"
      End
      Begin PRJFW_EDIT.TxtEdit TXT_TOTCARTE 
         Height          =   300
         Left            =   7545
         TabIndex        =   48
         ToolTipText     =   "PZ"
         Top             =   825
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
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL4 
         Height          =   300
         Left            =   90
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1875
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         Caption         =   "Tot. Spessori"
      End
      Begin PRJFW_EDIT.TxtEdit TXT_TOTSPESSORI 
         Height          =   300
         Left            =   1245
         TabIndex        =   46
         ToolTipText     =   "PZ"
         Top             =   1860
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
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL12 
         Height          =   300
         Left            =   3135
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1545
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Caption         =   "Tot. Carta ( Kg.)"
      End
      Begin PRJFW_EDIT.TxtEdit TXT_TOTKGCARTA 
         Height          =   300
         Left            =   4305
         TabIndex        =   30
         ToolTipText     =   "PZ"
         Top             =   1515
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
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL10 
         Height          =   300
         Left            =   3150
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1185
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         Caption         =   "Tubi per Bancale"
      End
      Begin PRJFW_EDIT.TxtEdit TXT_TUBXBANCALE 
         Height          =   300
         Left            =   4305
         TabIndex        =   28
         ToolTipText     =   "PZ"
         Top             =   1170
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
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL9 
         Height          =   300
         Left            =   90
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1515
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         Caption         =   "Lunghezza"
      End
      Begin PRJFW_EDIT.TxtEdit TXT_LUNGHEZZA 
         Height          =   300
         Left            =   1245
         TabIndex        =   26
         ToolTipText     =   "PZ"
         Top             =   1515
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
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL7 
         Height          =   300
         Left            =   75
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         Caption         =   "Diam. Medio"
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DMEDIO 
         Height          =   300
         Left            =   1245
         TabIndex        =   24
         ToolTipText     =   "PZ"
         Top             =   1170
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
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL6 
         Height          =   300
         Left            =   3150
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   825
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         Caption         =   "Diam. Esterno"
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESTERNO 
         Height          =   300
         Left            =   4320
         TabIndex        =   22
         ToolTipText     =   "PZ"
         Top             =   825
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
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL2 
         Height          =   300
         Left            =   75
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         Caption         =   "Diam. Interno"
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DINTERNO 
         Height          =   300
         Left            =   1245
         TabIndex        =   20
         ToolTipText     =   "PZ"
         Top             =   825
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
      Begin PRJFW_TmsLabel.TMS_LABEL LBL_QTA 
         Height          =   300
         Left            =   90
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   90
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         Caption         =   "Tipo quantità"
      End
      Begin PRJFW_COMBOBOX.TMS_COMBO CMB_TIPOQTA 
         Height          =   315
         Left            =   1260
         TabIndex        =   17
         Top             =   60
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
      Begin PRJFW_TmsLabel.TMS_LABEL LBL_UM 
         Height          =   300
         Left            =   3420
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   90
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   529
         Caption         =   "UM 1"
      End
      Begin PRJFW_EDIT.TxtEdit TXT_UM1 
         Height          =   300
         Left            =   3840
         TabIndex        =   15
         ToolTipText     =   "UM 1"
         Top             =   90
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
      Begin PRJFW_EDIT.TxtEdit TXT_FAM 
         Height          =   300
         Left            =   1440
         TabIndex        =   14
         ToolTipText     =   "Codice Famiglia"
         Top             =   480
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
      Begin PRJFW_EDIT.TxtEdit TXT_SFAM 
         Height          =   300
         Left            =   2055
         TabIndex        =   13
         ToolTipText     =   "Codice Sottofamiglia"
         Top             =   480
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
      Begin PRJFW_EDIT.TxtEdit TXT_GRUP 
         Height          =   300
         Left            =   2670
         TabIndex        =   12
         ToolTipText     =   "Codice Gruppo"
         Top             =   480
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
      Begin PRJFW_EDIT.TxtEdit TXT_SGRUP 
         Height          =   300
         Left            =   3285
         TabIndex        =   11
         ToolTipText     =   "Codice Sottogruppo"
         Top             =   480
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
      Begin PRJFW_TmsLabel.TMS_LABEL LBL_FAM 
         Height          =   300
         Left            =   90
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   529
         Caption         =   "Fm-Sfm-Gr-Sg"
      End
      Begin PRJFW_EDIT.TxtEdit TXT_PZ 
         Height          =   300
         Left            =   5460
         TabIndex        =   9
         ToolTipText     =   "PZ"
         Top             =   90
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
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL1 
         Height          =   300
         Left            =   5220
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         Caption         =   "PZ"
      End
      Begin PRJFW_EDIT.TxtEdit TXT_DESGRUSTAT1 
         Height          =   300
         Left            =   5460
         TabIndex        =   7
         ToolTipText     =   "PZ"
         Top             =   480
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
      Begin PRJFW_EDIT.TxtEdit TXT_GRST2 
         Height          =   300
         Left            =   4500
         TabIndex        =   6
         ToolTipText     =   "Codice Sottogruppo"
         Top             =   480
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
      Begin PRJFW_TmsLabel.TMS_LABEL TMS_LABEL3 
         Height          =   300
         Left            =   3900
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   480
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   529
         Caption         =   "Gr.St.2"
      End
   End
   Begin PRJFW_FRAME.TMS_FRAME TMS_FRAME2 
      Height          =   555
      Left            =   30
      Top             =   1050
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   979
      Begin PRJFW_EDIT.TxtEdit TXT_PADREDISTINTA 
         Height          =   300
         Left            =   90
         TabIndex        =   95
         Top             =   90
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   529
         Enabled         =   0   'False
         MaxChar         =   25
         Numerico        =   0   'False
         Carattere       =   0   'False
         IsDbField       =   0   'False
         MaxWidth        =   25
      End
   End
   Begin PRJFW_FRAME.TMS_FRAME TMS_FRAME4 
      Height          =   5310
      Left            =   555
      Top             =   3120
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   9366
      Begin PRJFW_GRIDNAV.TMS_GRIDNAV GRIDNAV 
         Height          =   360
         Left            =   12480
         TabIndex        =   33
         Top             =   4890
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   635
         Primo           =   0   'False
         Precedente      =   0   'False
         Successivo      =   0   'False
         Ultimo          =   0   'False
      End
      Begin PRJFW_EDIT.TxtEdit TXT_Articolo 
         Height          =   300
         Left            =   1065
         TabIndex        =   60
         Top             =   3615
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         MaxChar         =   25
         Numerico        =   0   'False
         Carattere       =   0   'False
         DBField         =   "ARTICOLO"
      End
      Begin PRJFW_EDIT.TxtEdit TXT_id 
         Height          =   300
         Left            =   315
         TabIndex        =   44
         Tag             =   "id"
         ToolTipText     =   "id"
         Top             =   1245
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   529
         MaxChar         =   4
         Numerico        =   0   'False
         Carattere       =   0   'False
         DBField         =   "id"
         Caption         =   "id"
         MaxWidth        =   4
      End
      Begin PRJFW_EDIT.TxtEdit TXT_Colla 
         Height          =   300
         Left            =   1890
         TabIndex        =   43
         Tag             =   "Colla"
         ToolTipText     =   "Colla"
         Top             =   1395
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   529
         MaxChar         =   20
         Numerico        =   0   'False
         Carattere       =   0   'False
         DBField         =   "Colla"
         Caption         =   "Colla"
         MaxWidth        =   20
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_PosizioneA 
         Height          =   300
         Left            =   945
         TabIndex        =   42
         Tag             =   "PosizioneA"
         Top             =   2085
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         DBField         =   "PosizioneA"
         Caption         =   "PosizioneA"
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_PosizioneDa 
         Height          =   300
         Left            =   825
         TabIndex        =   41
         Tag             =   "PosizioneDa"
         Top             =   930
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   529
         DBField         =   "PosizioneDa"
         Caption         =   "PosizioneDa"
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_KGPezzi 
         Height          =   300
         Left            =   6045
         TabIndex        =   40
         Tag             =   "KGPezzi"
         Top             =   1335
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   529
         DBField         =   "KGPezzi"
         Caption         =   "KGPezzi"
         MaxWidth        =   20
         MaxChar         =   24
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_Grammi_Metro 
         Height          =   300
         Left            =   7620
         TabIndex        =   39
         Tag             =   "Grammi_Metro"
         Top             =   555
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   529
         DBField         =   "Grammi_Metro"
         Caption         =   "Grammi_Metro"
         MaxWidth        =   20
         MaxChar         =   24
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_Grammatura 
         Height          =   300
         Left            =   2535
         TabIndex        =   38
         Tag             =   "Peso"
         Top             =   720
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   529
         DBField         =   "GRAMMATURA"
         Caption         =   "Peso"
         MaxWidth        =   20
         MaxChar         =   24
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_Spessore 
         Height          =   300
         Left            =   3555
         TabIndex        =   37
         Tag             =   "Spessore"
         Top             =   3405
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   529
         DBField         =   "Spessore"
         Caption         =   "Spessore"
         MaxWidth        =   20
         MaxChar         =   24
      End
      Begin PRJFW_EDITNUM.TxtEditNum TXT_Qta 
         Height          =   300
         Left            =   6660
         TabIndex        =   36
         Tag             =   "Qta"
         Top             =   2775
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   529
         DBField         =   "Qta"
         Caption         =   "Qta"
         MaxWidth        =   20
         MaxChar         =   24
      End
      Begin PRJFW_EDIT.TxtEdit TXT_Descrizione 
         Height          =   300
         Left            =   8820
         TabIndex        =   35
         Tag             =   "Descrizione"
         ToolTipText     =   "Descrizione"
         Top             =   1905
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   529
         MaxChar         =   72
         Numerico        =   0   'False
         Carattere       =   0   'False
         DBField         =   "Descrizione"
         Caption         =   "Descrizione"
         MaxWidth        =   20
      End
      Begin PRJFW_EDIT.TxtEdit TXT_ArticoloPadre 
         Height          =   300
         Left            =   1845
         TabIndex        =   34
         Tag             =   "ArticoloPadre"
         ToolTipText     =   "ArticoloPadre"
         Top             =   225
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   529
         MaxChar         =   25
         Numerico        =   0   'False
         Carattere       =   0   'False
         DBField         =   "ArticoloPadre"
         Caption         =   "ArticoloPadre"
         MaxWidth        =   20
      End
   End
   Begin MSComctlLib.ImageList IMGL_FASI_PF 
      Left            =   4605
      Top             =   1725
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMIT_CONFDIST.frx":0442
            Key             =   "UNSEL"
            Object.Tag             =   "0"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMIT_CONFDIST.frx":07DC
            Key             =   "SEL"
            Object.Tag             =   "1"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList IMGL_FASI_SL 
      Left            =   5745
      Top             =   1635
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMIT_CONFDIST.frx":0B76
            Key             =   "UNSEL"
            Object.Tag             =   "0"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMIT_CONFDIST.frx":0F10
            Key             =   "SEL"
            Object.Tag             =   "1"
         EndProperty
      EndProperty
   End
   Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON TMS_FLATBUTTON5 
      Height          =   765
      Left            =   2250
      TabIndex        =   99
      Top             =   1620
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   1349
      Caption         =   "A P R I  D I S T I N T A"
      ButtonBorderColor=   0
      ButtonForeColor =   16777215
      ButtonHilightBorderColor=   9473677
   End
   Begin PRJFW_FLATBUTTON.TMS_FLATBUTTON CMD_CREADISTINTA 
      Height          =   765
      Left            =   30
      TabIndex        =   32
      Top             =   1620
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   1349
      Caption         =   "C R E A  D I S T I N T A"
      ButtonBorderColor=   0
      ButtonForeColor =   16777215
      ButtonHilightBorderColor=   9473677
   End
End
Attribute VB_Name = "FRMIT_CONFDIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'to do

' media entrate / uscite
' campo rosso list base, se brezzo base non è maggiore di x % del costo

Option Explicit

Public Gcls_Global                              As CLSFW_Global
Public Gcls_Log                                 As CLSFW_SrvLog
Public Gcon_Connect                             As ADODB.Connection
Public Gcls_Connect                             As New CLSFW_SetConnect
Public Gstr_Connect                             As String

Public ActiveInterface                          As Cinterface
Public ActiveClass                              As CLSUO_CONFDIST
Private pbol_alreadyloaded                      As Boolean

Public Gcls_RecSet_SitGiacenze                  As New CLSFW_Recordset
Public WithEvents Grst_SitGiacenze              As ADODB.Recordset
Attribute Grst_SitGiacenze.VB_VarHelpID = -1
Public WithEvents FME_CCS_SKPROD                As CLSFW_VIRTUALFRAME
Attribute FME_CCS_SKPROD.VB_VarHelpID = -1
Public WithEvents FME_CONFDIST                  As CLSFW_VIRTUALFRAME
Attribute FME_CONFDIST.VB_VarHelpID = -1
Public WithEvents FME_IMBALLI                   As CLSFW_VIRTUALFRAME
Attribute FME_IMBALLI.VB_VarHelpID = -1
Public WithEvents FME_IMBALLIPF                   As CLSFW_VIRTUALFRAME
Attribute FME_IMBALLIPF.VB_VarHelpID = -1

Public WithEvents FME_FASIPF                   As CLSFW_VIRTUALFRAME
Attribute FME_FASIPF.VB_VarHelpID = -1
Public WithEvents FME_FASISL                   As CLSFW_VIRTUALFRAME
Attribute FME_FASISL.VB_VarHelpID = -1

Public Gstr_SQL_SitGiacenze                     As String

Public Gstr_DittaCorrente                       As String

Public Prst_Progressivi                         As ADODB.Recordset

'Enzo 200703 - Carico listini vendita e acquisto
Public Grst_RecSet_LI11VEN                    As ADODB.Recordset
Public Grst_RecSet_LI11_appendVEN             As ADODB.Recordset
Public Grst_RecSet_LI11ACQ                    As ADODB.Recordset
Public Grst_RecSet_LI11_appendACQ             As ADODB.Recordset
Public Grst_RecSet_LI11ACQ_TOT                As ADODB.Recordset
Public Grst_RecSet_LI11_appendACQ_TOT         As ADODB.Recordset

Private rstARTICOLI                           As ADODB.Recordset
Private rstCONFDIST                           As ADODB.Recordset
Private rstIMBALLI                          As ADODB.Recordset
Private rstIMBALLIPF                        As ADODB.Recordset
Private rstFASIPF                           As ADODB.Recordset
Private rstFASISL                           As ADODB.Recordset

Private WithEvents clsBOImport              As IEBO_IMPORTAZIONE.CLSIE_BOIMPORT
Attribute clsBOImport.VB_VarHelpID = -1
Private bolAnnullaImportazione              As Boolean

Private Gcls_RecordPadre                    As CLSFW_Recordset


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

Private Sub CMD_ANAGRAFICA_Click()

Set Cls_ConnectMagazzino.ActiveInterface = ActiveInterface
        
        Cls_ConnectMagazzino.Left = 10
        Cls_ConnectMagazzino.Top = 1000
        

                                                                
        Call Cls_ConnectMagazzino.ArticoloAnagrafica(RTrimN(TXT_CODART.Text))
        ActiveInterface.IsActive = True
        Set Cls_ConnectMagazzino.ActiveInterface = Nothing
        Set ActiveInterface.ClsGlobal.ActiveInterface = ActiveInterface
End Sub

Public Function CallDisba(ByVal WkAssieme As Variant, _
                          ByVal WkVarianteAssieme As Variant, _
                          ByVal WkTipoDistinta As Variant, _
                          ByVal WkCodiceDb As Variant) As Boolean
    '
    ' Trap degli errori
    '
    On Error GoTo Err_CallDisba
    '
    ' Disattivo il presente programma
    '
    ActiveInterface.IsActive = False
    '
    ' Setto i parametri alla classe connect
    '
    Set Pcls_Connect_Produzione.ActiveInterface = ActiveInterface
    Pcls_Connect_Produzione.CodiceAssieme = WkAssieme
    Pcls_Connect_Produzione.CodiceVarianteArticolo = WkVarianteAssieme
    Pcls_Connect_Produzione.TipoDistinta = WkTipoDistinta
    Pcls_Connect_Produzione.CodiceDistinta = WkCodiceDb
'    Cls_ConnectProduzione.TabellaDistintaBase = "PD22_DISBA"
'    Cls_ConnectProduzione.TabellaNoteDistintaBase = "PD23_NOTEDBISBA"
    Pcls_Connect_Produzione.CallDistintaBase
    '
    ' Attivo il presente programma
    '
    ActiveInterface.IsActive = True
    '
    ' Rilascio la chiamata
    '
    Set Pcls_Connect_Produzione.ActiveInterface = Nothing
    Pcls_Connect_Produzione.TerminateConnect
    '
    ' Riassegno l'interfaccia attiva
    '
    Set ActiveInterface.ClsGlobal.ActiveInterface = ActiveInterface
    '
    ' Esco
    '
    CallDisba = True
    Exit Function
Err_CallDisba:
    CallDisba = False
    Err.Clear
    Exit Function
End Function
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

Private Sub CMD_CALLDISBA_Click()
Call CallDisba(TXT_CODART.Text, "", 0, "")
End Sub

Private Sub CMD_CREADISTINTA_Click()

    TAB_DISTINTA.ActiveTab = 3

    If VerificaDatiCreazioneDistinta Then
    WriteTxtLog ("Dati Completi")

    'genera Distinta
        
            PreparaDatiGenerazioneDistinta
        
        Else
    'comunica anomalia
    WriteTxtLog ("Dati Incompleti")
    
    End If

End Sub


Private Sub WriteTxtLog(str As String)

TXT_LOG.Text = TXT_LOG.Text & str & vbCrLf

End Sub

Private Sub ImportazioneTracciato(NOME_TRACCIATO As String, NOME_FILE As String, CodiceStruttura As Integer, StrMsg As String)
    Dim strTestoPerLog      As String
    Dim LetturaTracciato    As String
On Error GoTo ErrTrap
    
    If NVL(NOME_TRACCIATO, "") = "" Then
        Exit Sub
    End If
    
    ProgressBar1.Visible = True
    ProgressBar1.Max = 1000
    
    bolAnnullaImportazione = True
    
    Set clsBOImport = New IEBO_IMPORTAZIONE.CLSIE_BOIMPORT
    With clsBOImport
        .CodiceStruttura = CodiceStruttura
        .NomeTracciato = NOME_TRACCIATO
        Set .connessione = Gcon_Connect
                
       ' .NomeFileSorgente = NOME_FILE
        
        Set .ActiveInterface = ActiveInterface
        
        .ImportaDati
        
        If .Stato <> 0 Then
            WriteTxtLog (.Errore & " in Importazione " & NOME_TRACCIATO)
        Else
            WriteTxtLog (NOME_TRACCIATO & ": " & StrMsg)
        End If
        
    End With
    Set clsBOImport = Nothing
    
    'ProgressBar1.Visible = False

Exit Sub
ErrTrap:
    Select Case VisualizzaErrore("ImportazioneVeraEPropria")
        Case vbAbort: Exit Sub
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Sub

Private Sub PreparaDatiGenerazioneDistinta()

' recupera struttura Padre - Figlio - Qta

Dim strSQL As String
Dim strSQLInsert As String
Dim oRsStruttura       As ADODB.Recordset
Dim Ripetizione        As Integer
Dim i As Integer
Dim ArticoloPF As String
Dim ArticoloSL As String
Dim QtaTubiPerBancale As Double
Dim FasePF As String
TXT_LOG.Text = ""
 WriteTxtLog ("INIZIO CREAZIONE DISTINTA " & ArticoloPF)

'pulisco tabella
Gcon_Connect.Execute "DELETE IT01_DISTINTA"

ArticoloSL = TXT_PADREDISTINTA.Text
ArticoloPF = TXT_CODART.Text
FasePF = GetValFromQuery("SELECT IT01_CODFASE FROM IT01_FASIPF WHERE (IT01_SEQ > 0) ORDER BY IT01_SEQ")


strSQLInsert = "INSERT INTO  IT01_DISTINTA ( IT01_PADRE, IT01_FIGLIO, IT01_QTA, IT01_FASE, IT01_FLGSIMILAV ) VALUES "
strSQLInsert = strSQLInsert & "('" & ArticoloPF & "','" & ArticoloSL & "',1,'" & FasePF & "',1)"
Gcon_Connect.Execute strSQLInsert
'caricamento Articolo sl
Call ImportazioneTracciato("IT01_ARTICOLOSL", "Log.log", 1, "Caricamento Articolo SL: " & ArticoloSL)

WriteTxtLog ("")

strSQL = "Select * from IT00_CONFDISTINTE WHERE (PosizioneDa > 0) order by PosizioneDa"
  
Set oRsStruttura = Gcon_Connect.Execute(strSQL, , adCmdText)
If Not oRsStruttura.EOF Then
      Do While Not oRsStruttura.EOF
       
            Ripetizione = (NVL(oRsStruttura("PosizioneA"), 0) - NVL(oRsStruttura("PosizioneDA"), 0)) + 1
            'If Ripetizione = 0 Then Ripetizione
      
            For i = 1 To Ripetizione
                
                strSQLInsert = "INSERT INTO  IT01_DISTINTA ( IT01_PADRE, IT01_FIGLIO, IT01_QTA, IT01_FASE ) VALUES "
                strSQLInsert = strSQLInsert & "('" & ArticoloSL & "','" & oRsStruttura("Articolo") & "'," & oRsStruttura("Qta") & ",'" & FasePF & "')"
                Gcon_Connect.Execute strSQLInsert
                
                WriteTxtLog ("Inserimento Legame > " & oRsStruttura("Articolo") & "-" & ArticoloSL)
            
            Next
            
      
        oRsStruttura.MoveNext
      Loop
End If
  
 
Set oRsStruttura = Nothing

strSQL = "Select * from  IT01_ARTICOLIIMBALLI WHERE (Quantita > 0) "

Set oRsStruttura = Gcon_Connect.Execute(strSQL, , adCmdText)
If Not oRsStruttura.EOF Then
      Do While Not oRsStruttura.EOF
       
                QtaTubiPerBancale = Format(oRsStruttura("Quantita") / CDbl(TXT_TUBXBANCALE.Text), "0.000000000000")
                strSQLInsert = "INSERT INTO  IT01_DISTINTA ( IT01_PADRE, IT01_FIGLIO, IT01_QTA, IT01_FASE ) VALUES "
                strSQLInsert = strSQLInsert & "('" & ArticoloSL & "','" & oRsStruttura("Articolo") & "'," & SQLDouble(QtaTubiPerBancale) & ",'" & FasePF & "')"
                Gcon_Connect.Execute strSQLInsert
            
        oRsStruttura.MoveNext
      Loop
End If
  
 
Set oRsStruttura = Nothing
WriteTxtLog ("Inserimento Imballo SL")

strSQL = "Select * from  IT01_ARTICOLIIMBALLI_PF WHERE (Quantita > 0) "

Set oRsStruttura = Gcon_Connect.Execute(strSQL, , adCmdText)
If Not oRsStruttura.EOF Then
      Do While Not oRsStruttura.EOF
       
                QtaTubiPerBancale = Format(oRsStruttura("Quantita") / CDbl(TXT_TUBXBANCALE.Text), "0.000000000000")
                strSQLInsert = "INSERT INTO  IT01_DISTINTA ( IT01_PADRE, IT01_FIGLIO, IT01_QTA, IT01_FASE ) VALUES "
                strSQLInsert = strSQLInsert & "('" & ArticoloPF & "','" & oRsStruttura("Articolo") & "'," & SQLDouble(QtaTubiPerBancale) & ",'" & FasePF & "')"
                Gcon_Connect.Execute strSQLInsert
            

            
      
        oRsStruttura.MoveNext
      Loop
End If
Set oRsStruttura = Nothing
WriteTxtLog ("Inserimento Imballo PF")


     strSQL = " IF NOT EXISTS (SELECT "
     strSQL = strSQL & "     * "
     strSQL = strSQL & "   FROM PD52_ANAGCICLI "
     strSQL = strSQL & "   WHERE PD52_ANAGCICLI.PD52_CODICE = '" & ArticoloSL & "')"
     strSQL = strSQL & "   INSERT INTO PD52_ANAGCICLI (PD52_DITTA_CG18, PD52_CODICE, PD52_DESCR, PD52_VERSIONE) "
     strSQL = strSQL & "    Values (" & Gstr_DittaCorrente & ",'" & ArticoloSL & "','Ciclo di " & ArticoloSL & "',0) "
     Gcon_Connect.Execute strSQL
     
     strSQL = " UPDATE PD18_ARTPROD "
     strSQL = strSQL & " SET PD18_CODCICLO = PD18_CODART_MG66 "
     strSQL = strSQL & " WHERE PD18_CODART_MG66 = '" & ArticoloSL & "'"
     strSQL = strSQL & " AND PD18_DITTA_CG18 = " & Gstr_DittaCorrente
     Gcon_Connect.Execute strSQL
     
          strSQL = " IF NOT EXISTS (SELECT "
     strSQL = strSQL & "     * "
     strSQL = strSQL & "   FROM PD52_ANAGCICLI "
     strSQL = strSQL & "   WHERE PD52_ANAGCICLI.PD52_CODICE = '" & ArticoloPF & "')"
     strSQL = strSQL & "   INSERT INTO PD52_ANAGCICLI (PD52_DITTA_CG18, PD52_CODICE, PD52_DESCR, PD52_VERSIONE) "
     strSQL = strSQL & "    Values (" & Gstr_DittaCorrente & ",'" & ArticoloPF & "','Ciclo di " & ArticoloPF & "',0) "
     Gcon_Connect.Execute strSQL
     
     strSQL = " UPDATE PD18_ARTPROD "
     strSQL = strSQL & " SET PD18_CODCICLO = PD18_CODART_MG66 "
     strSQL = strSQL & " WHERE PD18_CODART_MG66 = '" & ArticoloPF & "'"
     strSQL = strSQL & " AND PD18_DITTA_CG18 = " & Gstr_DittaCorrente
     Gcon_Connect.Execute strSQL
     
     
     'Ciclo PF
       Gcon_Connect.Execute "delete PD48_CICLI where PD48_DITTA_CG18 = " & Gstr_DittaCorrente & " AND PD48_CICLO_PD52 = '" & ArticoloPF & "' AND PD48_IDDISBA_PD95 IS NULL "

     strSQL = "Select * from  IT01_FASIPF WHERE        (IT01_SEQ > 0) ORDER BY IT01_SEQ "

Set oRsStruttura = Gcon_Connect.Execute(strSQL, , adCmdText)
If Not oRsStruttura.EOF Then
      Do While Not oRsStruttura.EOF
                'controllo esistenza per articolo e fase
                strSQLInsert = " INSERT INTO [dbo].[PD48_CICLI] "
                strSQLInsert = strSQLInsert & "            ([PD48_DITTA_CG18] "
                strSQLInsert = strSQLInsert & "            ,[PD48_CICLO_PD52] "
                strSQLInsert = strSQLInsert & "            ,[PD48_VERSIONE_PD52] "
                strSQLInsert = strSQLInsert & "            ,[PD48_SEQFASE] "
                strSQLInsert = strSQLInsert & "            ,[PD48_CODFASE_PD12] "
                strSQLInsert = strSQLInsert & "            ,[PD48_PROGFASE] "
                strSQLInsert = strSQLInsert & "            ,[PD48_INDTIPOPROV] "
                strSQLInsert = strSQLInsert & "            ,[PD48_IDDISBA_PD95] "
                strSQLInsert = strSQLInsert & "            ,[PD48_CODREP_PD07] "
                strSQLInsert = strSQLInsert & "            ,[PD48_CODFORN_CG44] "
                strSQLInsert = strSQLInsert & "            ,[PD48_MACCHINA_PD08] "
                strSQLInsert = strSQLInsert & "            ,[PD48_INDUMSETUP] "
                strSQLInsert = strSQLInsert & "            ,[PD48_INDUMLAV] "
                strSQLInsert = strSQLInsert & "            ,[PD48_INDUMTOTLAV] "
                strSQLInsert = strSQLInsert & "            ,[PD48_INDUMATTESA] "
                strSQLInsert = strSQLInsert & "            ,[PD48_INDUMCODA] "
                strSQLInsert = strSQLInsert & "  "
                strSQLInsert = strSQLInsert & "           ) "
                strSQLInsert = strSQLInsert & "      VALUES "
                strSQLInsert = strSQLInsert & "            (" & Gstr_DittaCorrente
                strSQLInsert = strSQLInsert & "            ,'" & ArticoloPF & "'"
                strSQLInsert = strSQLInsert & "            ,0 "
                strSQLInsert = strSQLInsert & "            ," & oRsStruttura("IT01_SEQ")
                strSQLInsert = strSQLInsert & "            ,'" & oRsStruttura("IT01_CODFASE") & "'"
                strSQLInsert = strSQLInsert & "            ,0"
                strSQLInsert = strSQLInsert & "            ,0"
                strSQLInsert = strSQLInsert & "            ,Null"
                strSQLInsert = strSQLInsert & "            ,'" & Trim(oRsStruttura("IT01_CODREP_PD07")) & "'"
                strSQLInsert = strSQLInsert & "            ,Null"
                strSQLInsert = strSQLInsert & "            ,'" & Trim(oRsStruttura("IT01_MACCHINA_PD08")) & "'"
                strSQLInsert = strSQLInsert & "            ,0"
                strSQLInsert = strSQLInsert & "            ,0"
                strSQLInsert = strSQLInsert & "            ,0"
                strSQLInsert = strSQLInsert & "            ,0"
                strSQLInsert = strSQLInsert & "            ,0"
                strSQLInsert = strSQLInsert & "            )"
                
                
                
                Gcon_Connect.Execute strSQLInsert
            

            
      
        oRsStruttura.MoveNext
      Loop
End If
Set oRsStruttura = Nothing
WriteTxtLog ("Inserimento Ciclo PF")

  'Ciclo SL
  
  Gcon_Connect.Execute "delete PD48_CICLI where PD48_DITTA_CG18 = " & Gstr_DittaCorrente & " AND PD48_CICLO_PD52 = '" & ArticoloSL & "' AND PD48_IDDISBA_PD95 IS NULL "
     strSQL = "Select * from  IT01_FASISL WHERE        (IT01_SEQ > 0) ORDER BY IT01_SEQ "

Set oRsStruttura = Gcon_Connect.Execute(strSQL, , adCmdText)
If Not oRsStruttura.EOF Then
      Do While Not oRsStruttura.EOF
                 'controllo esistenza per articolo e fase
                strSQLInsert = " INSERT INTO [dbo].[PD48_CICLI] "
                strSQLInsert = strSQLInsert & "            ([PD48_DITTA_CG18] "
                strSQLInsert = strSQLInsert & "            ,[PD48_CICLO_PD52] "
                strSQLInsert = strSQLInsert & "            ,[PD48_VERSIONE_PD52] "
                strSQLInsert = strSQLInsert & "            ,[PD48_SEQFASE] "
                strSQLInsert = strSQLInsert & "            ,[PD48_CODFASE_PD12] "
                strSQLInsert = strSQLInsert & "            ,[PD48_PROGFASE] "
                strSQLInsert = strSQLInsert & "            ,[PD48_INDTIPOPROV] "
                strSQLInsert = strSQLInsert & "            ,[PD48_IDDISBA_PD95] "
                strSQLInsert = strSQLInsert & "            ,[PD48_CODREP_PD07] "
                strSQLInsert = strSQLInsert & "            ,[PD48_CODFORN_CG44] "
                strSQLInsert = strSQLInsert & "            ,[PD48_MACCHINA_PD08] "
                strSQLInsert = strSQLInsert & "            ,[PD48_INDUMSETUP] "
                strSQLInsert = strSQLInsert & "            ,[PD48_INDUMLAV] "
                strSQLInsert = strSQLInsert & "            ,[PD48_INDUMTOTLAV] "
                strSQLInsert = strSQLInsert & "            ,[PD48_INDUMATTESA] "
                strSQLInsert = strSQLInsert & "            ,[PD48_INDUMCODA] "
                strSQLInsert = strSQLInsert & "  "
                strSQLInsert = strSQLInsert & "           ) "
                strSQLInsert = strSQLInsert & "      VALUES "
                strSQLInsert = strSQLInsert & "            (" & Gstr_DittaCorrente
                strSQLInsert = strSQLInsert & "            ,'" & ArticoloSL & "'"
                strSQLInsert = strSQLInsert & "            ,0 "
                strSQLInsert = strSQLInsert & "            ," & oRsStruttura("IT01_SEQ")
                strSQLInsert = strSQLInsert & "            ,'" & oRsStruttura("IT01_CODFASE") & "'"
                strSQLInsert = strSQLInsert & "            ,0"
                strSQLInsert = strSQLInsert & "            ,0"
                strSQLInsert = strSQLInsert & "            ,Null"
                strSQLInsert = strSQLInsert & "            ,'" & Trim(oRsStruttura("IT01_CODREP_PD07")) & "'"
                strSQLInsert = strSQLInsert & "            ,Null"
                strSQLInsert = strSQLInsert & "            ,'" & Trim(oRsStruttura("IT01_MACCHINA_PD08")) & "'"
                strSQLInsert = strSQLInsert & "            ,0"
                strSQLInsert = strSQLInsert & "            ,0"
                strSQLInsert = strSQLInsert & "            ,0"
                strSQLInsert = strSQLInsert & "            ,0"
                strSQLInsert = strSQLInsert & "            ,0"
                strSQLInsert = strSQLInsert & "            )"
                
                
                
                Gcon_Connect.Execute strSQLInsert
            

            
      
        oRsStruttura.MoveNext
      Loop
End If
Set oRsStruttura = Nothing
WriteTxtLog ("Inserimento Ciclo SL")

'caricamento distinta
Call ImportazioneTracciato("IT01_IMPDIBA", "Log.log", 8, "Importazione Distinta")
WriteTxtLog ("FINE CREAZIONE DISTINTA " & ArticoloPF)

End Sub


Private Function VerificaDatiCreazioneDistinta() As Boolean
WriteTxtLog ("Inizia Verifica Dati")
On Error GoTo Errore:
    
    Dim strSQL As String
    
    If GetNumRecordFromQuery("SELECT * From IT00_CONFDISTINTE Where (PosizioneDa > 0)") = 0 Then
        VerificaDatiCreazioneDistinta = False
        Exit Function
    End If
    
    
    If GetNumRecordFromQuery("SELECT * From  IT01_ARTICOLIIMBALLI Where (Quantita > 0)") = 0 Then
        VerificaDatiCreazioneDistinta = False
        Exit Function
    End If
    '
    
    If GetNumRecordFromQuery("SELECT * From  IT01_ARTICOLIIMBALLI_PF Where (Quantita > 0)") = 0 Then
        VerificaDatiCreazioneDistinta = False
        Exit Function
    End If
    
    
    If GetNumRecordFromQuery("SELECT * From  IT01_FASISL Where (IT01_SEQ > 0)") = 0 Then
        VerificaDatiCreazioneDistinta = False
        Exit Function
    End If
    
    If GetNumRecordFromQuery("SELECT * From  IT01_FASIPF Where (IT01_SEQ > 0)") = 0 Then
        VerificaDatiCreazioneDistinta = False
        Exit Function
    End If
    
    
    VerificaDatiCreazioneDistinta = True
Exit Function

Errore:
    VerificaDatiCreazioneDistinta = False

End Function



Private Function GetNumRecordFromQuery(SQL As String)
'
  On Error GoTo ErrTrap
  
  Dim MyRst       As ADODB.Recordset
  

  
  Set MyRst = Gcon_Connect.Execute(SQL, , adCmdText)
  If Not MyRst.EOF Then
      GetNumRecordFromQuery = MyRst.RecordCount
  Else
      GetNumRecordFromQuery = 0
  End If
  
 
  Set MyRst = Nothing

  
  Exit Function
ErrTrap:
 GetNumRecordFromQuery = 0

End Function

Private Sub CMD_ELABORA_ButtonClick()
On Error GoTo Err
    Decimali
    TAB_DISTINTA.ActiveTab = 0
    If RTrimN(TXT_CODART.Text) <> "" And TXT_CODART.IsValid Then
        TXT_CODART.Enabled = False
        CMB_TIPOQTA.Enabled = False
        'CMD_ELABORA.Enabled = False
    Else
        MsgBox "Campo obbligatorio mancante!", vbCritical, "Informazione"
        TXT_CODART.SetTextFocus
        Exit Sub
    End If
    
TXT_DINTERNO.Text = "0" 'GetValFromQuery("select PD27_VARIABILE_X from pd27_ANAFORMULE where Pd27_DITTA_CG18 = " & Gstr_DittaCorrente & " and PD27_CODART_MG66 = '" & TXT_CODART.Text & "'")
TXT_DESTERNO.Text = ""
TXT_DMEDIO.Text = GetValFromQuery("select PD27_VARIABILE_z from pd27_ANAFORMULE where Pd27_DITTA_CG18 = " & Gstr_DittaCorrente & " and PD27_CODART_MG66 = '" & TXT_CODART.Text & "'")
TXT_LUNGHEZZA.Text = GetValFromQuery("select PD27_VARIABILE_Y from pd27_ANAFORMULE where Pd27_DITTA_CG18 = " & Gstr_DittaCorrente & " and PD27_CODART_MG66 = '" & TXT_CODART.Text & "'")
TXT_TUBXBANCALE.Text = GetValFromQuery("SELECT MG68_COLLIXBANCALE FROM MG68_CONFART WHERE (MG68_DITTA_CG18 = " & Gstr_DittaCorrente & ") AND (MG68_CODART_MG66 = '" & TXT_CODART.Text & "') AND (MG68_CODCONFEZ_MG96 = 'CD')")
TXT_TOTKGCARTA.Text = ""
TXT_PADREDISTINTA.Text = TXT_CODART.Text + "-SL"
'Call CaricaGrigliaArticoli


'
If Not TXT_TUBXBANCALE.IsValid Or NVL(TXT_TUBXBANCALE.Text, "") = "" Then

    MsgBox "Anagrafica incompleta"
    Exit Sub

End If


If Not TXT_PADREDISTINTA.IsValid Or NVL(TXT_PADREDISTINTA.Text, "") = "" Then

    MsgBox "Anagrafica incompleta"
    Exit Sub

End If

If Not TXT_LUNGHEZZA.IsValid Or NVL(TXT_LUNGHEZZA.Text, "") = "" Then

    MsgBox "Anagrafica incompleta"
    Exit Sub

End If



Call ImpostaVirtualFrame


    Call INITGRID_ARTICOLI
    
    'rstCONFDIST.MoveFirst
    Set QGRID_ARTICOLI.DataSource = rstCONFDIST
            
    QGRID_ARTICOLI.Refresh
            
    QGRID_ARTICOLI.Visible = True
    
Call ImpostaVirtualFrameImballi
    
    Call INITGRID_IMBALLI
    
    Set QGRID_IMBALLI.DataSource = rstIMBALLI
            
    QGRID_IMBALLI.Refresh
            
    QGRID_IMBALLI.Visible = True

Call ImpostaVirtualFrameImballiPF
    
    Call INITGRID_IMBALLI_PF
    
    Set QGRID_IMBALLIPF.DataSource = rstIMBALLIPF
            
    QGRID_IMBALLIPF.Refresh
            
    QGRID_IMBALLIPF.Visible = True
    
Call ImpostaVirtualFrameFasiPF
    
    Call INITGRID_FASI_PF
    
    Set QGRID_FASIPF.DataSource = rstFASIPF
            
    QGRID_FASIPF.Refresh
            
    QGRID_FASIPF.Visible = True
    
Call ImpostaVirtualFrameFasiSL
    
    Call INITGRID_FASI_SL
    
    Set QGRID_FASISL.DataSource = rstFASISL
            
    QGRID_FASISL.Refresh
            
    QGRID_FASISL.Visible = True
    
TAB_DISTINTA.ActiveTab = 0
    
Exit Sub

Err:
  Set Gcls_Log.vbError = Err
  Set Gcls_Log.ADOError = Gcon_Connect.Errors
  If Gcls_Log.ViewRunTimeErr("0_0_0_0", "Elabora", "Elabora.CMD_ELABORA_ButtonClick") = 1 Then
     Unload Me
  Else
     Resume Next
  End If
   

End Sub


Private Sub ImpostaVirtualFrameImballi()
Dim StringaSQL As String

    On Error GoTo ErrTrap

    'SQL string
    StringaSQL = " SELECT Articolo, DescrizioneArticolo, PesoN, Quantita, NumeroTubi, PesoUmido, PesoAsciutto FROM IT01_ARTICOLIIMBALLI "

    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstIMBALLI = Gcls_RecordPadre.Gpr_GetADORecord
    With rstIMBALLI
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With

    Set FME_IMBALLI = New FWBO_VIRTUALFRAME.CLSFW_VIRTUALFRAME
    'Virtual Frame Initialization
    FME_IMBALLI.Initialize ActiveInterface, Gcon_Connect, rstIMBALLI, StringaSQL, "", GRIDNAVIMBALLI, QGRID_IMBALLI

    ' binding components
   
   
   FME_IMBALLI.AddControl TXT_articoloimballi
   FME_IMBALLI.AddControl TXT_DescrizioneArticoloImballi
   FME_IMBALLI.AddControl TXT_PesoNImballi
   FME_IMBALLI.AddControl TXT_QuantitaImballi
   FME_IMBALLI.AddControl TXT_NumeroTubiImballi
   FME_IMBALLI.AddControl TXT_PesoUmidoImballi
   FME_IMBALLI.AddControl TXT_PesoAsciuttoImballi


   FME_IMBALLI.AddKey TXT_articoloimballi

    ' Controller initialization
    Set GRIDNAVIMBALLI.ActiveDll = ActiveInterface
    Set GRIDNAVIMBALLI.ActiveFrame = FME_IMBALLI

    ' GRIDNAV button
    GRIDNAVIMBALLI.Indietro = False
    GRIDNAVIMBALLI.Avanti = False
    GRIDNAVIMBALLI.Apri = False

      If rstIMBALLI.RecordCount > 0 Then
          FME_IMBALLI.Status = ActiveInterface.ProgramMode
      Else
          FME_IMBALLI.Status = tsInsert
      End If
      
      
      FME_IMBALLI.MsgOnUpdate = False

    Exit Sub
ErrTrap:
    Select Case VisualizzaErrore("ImpostaVirtualFrameImballi")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
End Sub


Private Sub ImpostaVirtualFrameImballiPF()
Dim StringaSQL As String

    On Error GoTo ErrTrap

    'SQL string
    StringaSQL = " SELECT Articolo, DescrizioneArticolo, PesoN, Quantita, NumeroTubi, PesoUmido, PesoAsciutto FROM IT01_ARTICOLIIMBALLI_PF "

    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstIMBALLIPF = Gcls_RecordPadre.Gpr_GetADORecord
    With rstIMBALLIPF
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With

    Set FME_IMBALLIPF = New FWBO_VIRTUALFRAME.CLSFW_VIRTUALFRAME
    'Virtual Frame Initialization
    FME_IMBALLIPF.Initialize ActiveInterface, Gcon_Connect, rstIMBALLIPF, StringaSQL, "", GRIDNAVIMBALLIPF, QGRID_IMBALLIPF

    ' binding components
   
   
   FME_IMBALLIPF.AddControl TXT_articoloimballiPF
   FME_IMBALLIPF.AddControl TXT_DescrizioneArticoloImballiPF
   FME_IMBALLIPF.AddControl TXT_PesoNImballiPF
   FME_IMBALLIPF.AddControl TXT_QuantitaImballiPF
   FME_IMBALLIPF.AddControl TXT_NumeroTubiImballiPF
   FME_IMBALLIPF.AddControl TXT_PesoUmidoImballiPF
   FME_IMBALLIPF.AddControl TXT_PesoAsciuttoImballiPF


   FME_IMBALLIPF.AddKey TXT_articoloimballiPF

    ' Controller initialization
    Set GRIDNAVIMBALLIPF.ActiveDll = ActiveInterface
    Set GRIDNAVIMBALLIPF.ActiveFrame = FME_IMBALLIPF

    ' GRIDNAV button
    GRIDNAVIMBALLIPF.Indietro = False
    GRIDNAVIMBALLIPF.Avanti = False
    GRIDNAVIMBALLIPF.Apri = False

      If rstIMBALLIPF.RecordCount > 0 Then
          FME_IMBALLIPF.Status = ActiveInterface.ProgramMode
      Else
          FME_IMBALLIPF.Status = tsInsert
      End If
      
      
      FME_IMBALLIPF.MsgOnUpdate = False

    Exit Sub
ErrTrap:
    Select Case VisualizzaErrore("ImpostaVirtualFrameImballiPF")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
End Sub


Private Sub ImpostaVirtualFrameFasiPF()
Dim StringaSQL As String

    On Error GoTo ErrTrap

    'SQL string
    StringaSQL = " SELECT IT01_CODFASE, IT01_SEQ, IT01_DESCR, IT01_CODREP_PD07, IT01_MACCHINA_PD08 FROM         IT01_FASIPF "

    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstFASIPF = Gcls_RecordPadre.Gpr_GetADORecord
    With rstFASIPF
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With

    Set FME_FASIPF = New FWBO_VIRTUALFRAME.CLSFW_VIRTUALFRAME
    'Virtual Frame Initialization
    FME_FASIPF.Initialize ActiveInterface, Gcon_Connect, rstFASIPF, StringaSQL, "", GRIDNAVFASIPF, QGRID_FASIPF

    ' binding components
   
   
   FME_FASIPF.AddControl TXT_SEQPF
   FME_FASIPF.AddControl TXT_DESCRFASEPF
   FME_FASIPF.AddControl TXT_REPFASEPF
   FME_FASIPF.AddControl TXT_FASEPF
   FME_FASIPF.AddControl TXT_MACCHINA_FPF
   


   FME_FASIPF.AddKey TXT_FASEPF

    ' Controller initialization
    Set GRIDNAVFASIPF.ActiveDll = ActiveInterface
    Set GRIDNAVFASIPF.ActiveFrame = FME_FASIPF

    ' GRIDNAV button
    GRIDNAVFASIPF.Indietro = False
    GRIDNAVFASIPF.Avanti = False
    GRIDNAVFASIPF.Apri = False

      If rstFASIPF.RecordCount > 0 Then
          FME_FASIPF.Status = ActiveInterface.ProgramMode
      Else
          FME_FASIPF.Status = tsInsert
      End If
      
      
      FME_FASIPF.MsgOnUpdate = False

    Exit Sub
ErrTrap:
    Select Case VisualizzaErrore("ImpostaVirtualFrameFasiPF")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
End Sub


Private Sub ImpostaVirtualFrameFasiSL()
Dim StringaSQL As String

    On Error GoTo ErrTrap

    'SQL string
    StringaSQL = "  SELECT IT01_CODFASE, IT01_SEQ, IT01_DESCR, IT01_CODREP_PD07, IT01_MACCHINA_PD08 FROM         IT01_FASISL  "

    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstFASISL = Gcls_RecordPadre.Gpr_GetADORecord
    With rstFASISL
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With

    Set FME_FASISL = New FWBO_VIRTUALFRAME.CLSFW_VIRTUALFRAME
    'Virtual Frame Initialization
    FME_FASISL.Initialize ActiveInterface, Gcon_Connect, rstFASISL, StringaSQL, "", GRIDNAVFASISL, QGRID_FASISL

    ' binding components
   
   
   FME_FASISL.AddControl TXT_SEQSL
   FME_FASISL.AddControl TXT_DESCRFASESL
   FME_FASISL.AddControl TXT_REPFASESL
   FME_FASISL.AddControl TXT_FASESL
   FME_FASISL.AddControl TXT_MACCHINA_FSL


   FME_FASISL.AddKey TXT_FASESL

    ' Controller initialization
    Set GRIDNAVFASISL.ActiveDll = ActiveInterface
    Set GRIDNAVFASISL.ActiveFrame = FME_FASISL

    ' GRIDNAV button
    GRIDNAVFASISL.Indietro = False
    GRIDNAVFASISL.Avanti = False
    GRIDNAVFASISL.Apri = False

      If rstFASISL.RecordCount > 0 Then
          FME_FASISL.Status = ActiveInterface.ProgramMode
      Else
          FME_FASISL.Status = tsInsert
      End If
      
      
      FME_FASISL.MsgOnUpdate = False

    Exit Sub
ErrTrap:
    Select Case VisualizzaErrore("ImpostaVirtualFrameFasiPF")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
End Sub

Private Sub FME_IMBALLI_ChangeStatus(ByVal fenm_status As FWBO_VIRTUALFRAME.EnumStatus, ByVal fenm_reason As FWBO_VIRTUALFRAME.EnumReason)
On Error GoTo ErrTrap

    Select Case fenm_status
    Case tsInsert
      TXT_articoloimballi.Enabled = True
      TXT_articoloimballi.SetFocus
    Case tsModify
      TXT_articoloimballi.Enabled = False
    End Select
    
 
    
    'Call AggiornaTotali
    
    
    'Call ImpostaVirtualFrame
    'rstCONFDIST.AbsolutePosition
    
    
Exit Sub
ErrTrap:
    Select Case VisualizzaErrore("FME_IMBALLI_ChangeStatus")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
End Sub

Private Sub FME_IMBALLIPF_ChangeStatus(ByVal fenm_status As FWBO_VIRTUALFRAME.EnumStatus, ByVal fenm_reason As FWBO_VIRTUALFRAME.EnumReason)
On Error GoTo ErrTrap

    Select Case fenm_status
    Case tsInsert
      TXT_articoloimballiPF.Enabled = True
      TXT_articoloimballiPF.SetFocus
    Case tsModify
      TXT_articoloimballiPF.Enabled = False
    End Select
    
 
    
    'Call AggiornaTotali
    
    
    'Call ImpostaVirtualFrame
    'rstCONFDIST.AbsolutePosition
    
    
Exit Sub
ErrTrap:
    Select Case VisualizzaErrore("FME_IMBALLI_PFChangeStatus")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
End Sub

Private Sub ImpostaVirtualFrame()
Dim StringaSQL As String

    On Error GoTo ErrTrap

    'SQL string
    StringaSQL = "SELECT id, ArticoloPadre, Articolo, Descrizione, Qta, Spessore, Grammatura, Grammi_Metro, KGPezzi, PosizioneDa, PosizioneA, Colla, SommaSpessori " & _
                 " FROM IT00_CONFDISTINTE "

    Set Gcls_RecordPadre = New CLSFW_Recordset
    Set rstCONFDIST = Gcls_RecordPadre.Gpr_GetADORecord
    With rstCONFDIST
        Set .ActiveConnection = Gcon_Connect
        .Source = StringaSQL
        .Open
        .MarshalOptions = adMarshalModifiedOnly
    End With

    Set FME_CONFDIST = New FWBO_VIRTUALFRAME.CLSFW_VIRTUALFRAME
    'Virtual Frame Initialization
    FME_CONFDIST.Initialize ActiveInterface, Gcon_Connect, rstCONFDIST, StringaSQL, "", GRIDNAV, QGRID_ARTICOLI

    ' binding components
   'FME_CONFDIST.AddControl TXT_id
   FME_CONFDIST.AddControl TXT_ArticoloPadre
   FME_CONFDIST.AddControl TXT_Articolo
   FME_CONFDIST.AddControl TXT_Descrizione
   FME_CONFDIST.AddControl TXT_Qta
   FME_CONFDIST.AddControl TXT_Spessore
   FME_CONFDIST.AddControl TXT_Grammatura
   FME_CONFDIST.AddControl TXT_Grammi_Metro
   FME_CONFDIST.AddControl TXT_KGPezzi
   FME_CONFDIST.AddControl TXT_PosizioneDa
   FME_CONFDIST.AddControl TXT_PosizioneA
   FME_CONFDIST.AddControl TXT_Colla

   FME_CONFDIST.AddKey TXT_Articolo

    ' Controller initialization
    Set GRIDNAV.ActiveDll = ActiveInterface
    Set GRIDNAV.ActiveFrame = FME_CONFDIST

    ' GRIDNAV button
    GRIDNAV.Indietro = False
    GRIDNAV.Avanti = False
    GRIDNAV.Apri = False

      If rstCONFDIST.RecordCount > 0 Then
          FME_CONFDIST.Status = ActiveInterface.ProgramMode
      Else
          FME_CONFDIST.Status = tsInsert
      End If
      
      
      FME_CONFDIST.MsgOnUpdate = False

    Exit Sub
ErrTrap:
    Select Case VisualizzaErrore("ImpostaVirtualFrame")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
End Sub

Private Sub FME_CONFDIST_BeforeMove(fbol_Cancel As Boolean, ByVal fenm_reason As FWBO_VIRTUALFRAME.EnumReason)
'If QGRID_ARTICOLI.DataSource("qta").Value > 0 Then
'    QGRID_ARTICOLI.DataSource("Grammi_Metro").Value = (CDbl(TXT_DMEDIO.Text) * 3.14) * (CDbl(TXT_TOTKGCARTA.Text) / 1000) * CDbl(QGRID_ARTICOLI.DataSource("qta").Value)
'End If

End Sub

Private Sub FME_CONFDIST_ChangeStatus(ByVal fenm_status As FWBO_VIRTUALFRAME.EnumStatus, ByVal fenm_reason As FWBO_VIRTUALFRAME.EnumReason)
On Error GoTo ErrTrap

    Select Case fenm_status
    Case tsInsert
      TXT_id.Enabled = True
      TXT_id.SetFocus
    Case tsModify
      TXT_id.Enabled = False
    End Select
    
 
    
    Call AggiornaTotali
    
    
    'Call ImpostaVirtualFrame
    'rstCONFDIST.AbsolutePosition
    
    
Exit Sub
ErrTrap:
    Select Case VisualizzaErrore("FME_CONFDIST_ChangeStatus")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
End Sub


Private Sub CaricaGrigliaArticoli()

    Dim strSQL As String
    Dim oRsTmpARTICOLI As ADODB.Recordset

    strSQL = "SELECT * FROM VMG66_ANAGARTBASE WHERE (MG66_FAM_MG53 = 'CART' OR  MG66_FAM_MG53 = 'COPE') "
    Set oRsTmpARTICOLI = Gcon_Connect.Execute(strSQL, , adCmdText)
    
    Call Carica_rstArticoli(oRsTmpARTICOLI)
    
    
    Call INITGRID_ARTICOLI
    
    rstARTICOLI.MoveFirst
    Set QGRID_ARTICOLI.DataSource = rstARTICOLI
            
    QGRID_ARTICOLI.Refresh
            
    QGRID_ARTICOLI.Visible = True


End Sub


Private Sub Carica_rstArticoli(TmpRecordSet As ADODB.Recordset)
    On Error GoTo ErrTrap
    Dim i As Integer

    Set rstARTICOLI = New ADODB.Recordset
    
    For i = 0 To TmpRecordSet.Fields.Count - 1
        rstARTICOLI.Fields.Append TmpRecordSet.Fields(i).Name, adVariant, TmpRecordSet.Fields(i).DefinedSize, adFldIsNullable
    Next
    rstARTICOLI.Fields.Append "Selezionato", adDecimal, 1, adFldIsNullable
    rstARTICOLI.Fields.Append "Qta", adDecimal, 10, adFldIsNullable
    rstARTICOLI.Fields.Append "PosizioneDa", adDecimal, 10, adFldIsNullable
    rstARTICOLI.Fields.Append "PosizioneA", adDecimal, 10, adFldIsNullable
    'rstARTICOLI.Fields.Append "Spessore", adDecimal, 10, adFldIsNullable
    'rstARTICOLI.Fields.Append "Grammatura", adDecimal, 10, adFldIsNullable
    rstARTICOLI.Fields.Append "Grammi_Metro", adDecimal, 10, adFldIsNullable
    rstARTICOLI.Fields.Append "KG_Pezzo", adDecimal, 10, adFldIsNullable
    rstARTICOLI.Fields.Append "Colla", adDecimal, 10, adFldIsNullable
    rstARTICOLI.CursorLocation = adUseClient
    rstARTICOLI.LockType = adLockOptimistic
    rstARTICOLI.CursorType = adOpenDynamic
    rstARTICOLI.Open
    
    If Not TmpRecordSet.BOF Then
        TmpRecordSet.MoveFirst
    End If
    
    While Not TmpRecordSet.EOF
        rstARTICOLI.AddNew
        For i = 0 To TmpRecordSet.Fields.Count - 1
            rstARTICOLI.Fields(TmpRecordSet.Fields(i).Name).Value = TmpRecordSet.Fields(i).Value
        Next
        
        rstARTICOLI.Fields!Selezionato = False
        rstARTICOLI.Fields!PosizioneDa = 0
        rstARTICOLI.Fields!PosizioneA = 0
        rstARTICOLI.Update
        TmpRecordSet.MoveNext
    Wend
    
    
   
    
Exit Sub

ErrTrap:
    Select Case VisualizzaErrore("Carica_rstArticoli")
        Case vbAbort
            Exit Sub
        Case vbRetry
            Resume
        Case vbIgnore
            Resume Next
    End Select
End Sub

Private Function GetValFromQuery(SQL As String)
'
  On Error GoTo ErrTrap
  
  Dim MyRst       As ADODB.Recordset
  

  
  Set MyRst = Gcon_Connect.Execute(SQL, , adCmdText)
  If Not MyRst.EOF Then
      GetValFromQuery = NVL(MyRst.Fields(0).Value, "")
  Else
      GetValFromQuery = ""
  End If
  
 
  Set MyRst = Nothing

  
  Exit Function
ErrTrap:
 GetValFromQuery = ""

End Function

Public Function NVL(Valore As Variant, ValIfNull As Variant) As Variant
    On Error Resume Next

    If IsEmpty(Valore) Or IsNull(Valore) Then
        NVL = ValIfNull
    Else
        If Trim(CStr(Valore)) = "" Then
            NVL = ValIfNull
        Else
            NVL = Trim(Valore)
        End If
    End If
    
    Err.Clear
End Function




Private Sub INITGRID_IMBALLI()
    Dim strSQL As String
    Dim ors                As ADODB.Recordset
    
    Dim cInit As New TMS_QGRID.CLSFW_INITQGRID
    Set cInit.ActiveInterface = ActiveInterface
    Set cInit.RSColumns = Nothing
    cInit.KeyField = "Articolo"
    cInit.CreateColumnsFromRS = False
    Set QGRID_IMBALLI.InitializationClass = cInit
    With QGRID_IMBALLI
        .CustomDrawCellEnabled = True
        .DataFormatEnabled = True

'id, ArticoloPadre, Articolo, Descrizione, Qta, Spessore, Grammatura, Grammi_Metro, KGPezzi, PosizioneDa, PosizioneA, Colla
        '.INIT_ADDColumnSELECTIONBOXEXT "SELEZIONATO", "Sel.", gedTextEdit, 0, 0, , True, True
        .INIT_ADDColumn "Articolo", "Articolo", gedTextEdit, 2000, True
        .INIT_ADDColumn "DescrizioneArticolo", "Descrizione Articolo", gedTextEdit, 3500, True
        .INIT_ADDColumn "PesoN", "Peso N", gedTextEdit, 1000, True
        .INIT_ADDColumn "Quantita", "Quantita", gedTextEdit, 1000, True
        .INIT_ADDColumn "NumeroTubi", "Numero Tubi", gedTextEdit, 800, True
        
        .INIT_ADDColumn "PesoUmido", "Peso Umido", gedTextEdit, 1000, True
        
        
        .INIT_ADDColumn "PesoAsciutto", "Peso Asciutto", gedTextEdit, 1000, True
                       
        .InitializeSTART
                 .EnableEditing = True
                 .MODCOL_Editing "Quantita", Enable
        .InitializeEND

    End With
    

    
    
    
    Set cInit = Nothing
End Sub

Private Sub INITGRID_IMBALLI_PF()
    Dim strSQL As String
    Dim ors                As ADODB.Recordset
    
    Dim cInit As New TMS_QGRID.CLSFW_INITQGRID
    Set cInit.ActiveInterface = ActiveInterface
    Set cInit.RSColumns = Nothing
    cInit.KeyField = "Articolo"
    cInit.CreateColumnsFromRS = False
    Set QGRID_IMBALLIPF.InitializationClass = cInit
    With QGRID_IMBALLIPF
        .CustomDrawCellEnabled = True
        .DataFormatEnabled = True

'id, ArticoloPadre, Articolo, Descrizione, Qta, Spessore, Grammatura, Grammi_Metro, KGPezzi, PosizioneDa, PosizioneA, Colla
        '.INIT_ADDColumnSELECTIONBOXEXT "SELEZIONATO", "Sel.", gedTextEdit, 0, 0, , True, True
        .INIT_ADDColumn "Articolo", "Articolo", gedTextEdit, 2000, True
        .INIT_ADDColumn "DescrizioneArticolo", "Descrizione Articolo", gedTextEdit, 3500, True
        .INIT_ADDColumn "PesoN", "Peso N", gedTextEdit, 1000, True
        .INIT_ADDColumn "Quantita", "Quantita", gedTextEdit, 1000, True
        .INIT_ADDColumn "NumeroTubi", "Numero Tubi", gedTextEdit, 800, True
        
        .INIT_ADDColumn "PesoUmido", "Peso Umido", gedTextEdit, 1000, True
        
        
        .INIT_ADDColumn "PesoAsciutto", "Peso Asciutto", gedTextEdit, 1000, True
                       
        .InitializeSTART
                 .EnableEditing = True
                 .MODCOL_Editing "Quantita", Enable
        .InitializeEND

    End With
    

    
    
    
    Set cInit = Nothing
End Sub

Private Sub INITGRID_FASI_PF()
    Dim strSQL As String
    Dim ors                As ADODB.Recordset
    
    Dim cInit As New TMS_QGRID.CLSFW_INITQGRID
    Set cInit.ActiveInterface = ActiveInterface
    Set cInit.RSColumns = Nothing
    cInit.KeyField = "IT01_CODFASE"
    cInit.CreateColumnsFromRS = False
    Set QGRID_FASIPF.InitializationClass = cInit
    With QGRID_FASIPF
        .CustomDrawCellEnabled = True
        .DataFormatEnabled = True

'id, ArticoloPadre, Articolo, Descrizione, Qta, Spessore, Grammatura, Grammi_Metro, KGPezzi, PosizioneDa, PosizioneA, Colla
        '.INIT_ADDColumnSELECTIONBOXEXT "SELEZIONATO", "Sel.", gedTextEdit, 0, 0, , True, True
        .INIT_ADDColumn "IT01_SEQ", "Sequenza", gedTextEdit, 800, True
        .INIT_ADDColumn "IT01_CODFASE", "Fase", gedTextEdit, 1000, True
        .INIT_ADDColumn "IT01_DESCR", "Descrizione Fase", gedTextEdit, 3500, True
        .INIT_ADDColumn "IT01_CODREP_PD07", "Reparto", gedTextEdit, 750, True
        .INIT_ADDColumn "IT01_MACCHINA_PD08", "Macchina", gedTextEdit, 750, True
        '.INIT_ADDColumnLookup_QUERY "IT01_MACCHINA_PD08", "SELECT PD08_MACCHINA,PD08_DESCR FROM PD08_MACCHINA", "Macchina", 1000, True
        '.INIT_ADDColumnLookup_VALUES "IT01_MACCHINA_PD08", xftString, "a,b,c"
        '.INIT_ADDColumn "IT01_MACCHINA_PD08", "Descrizione Fase", gedTextEdit, 3500, True
                               
        .InitializeSTART
                 .EnableEditing = True
                 .MODCOL_Editing "IT01_SEQ", Enable
        .InitializeEND

    End With
    

    
    
    
    Set cInit = Nothing
End Sub


Private Sub INITGRID_FASI_SL()
    Dim strSQL As String
    Dim ors                As ADODB.Recordset
    
    Dim cInit As New TMS_QGRID.CLSFW_INITQGRID
    Set cInit.ActiveInterface = ActiveInterface
    Set cInit.RSColumns = Nothing
    cInit.KeyField = "IT01_CODFASE"
    cInit.CreateColumnsFromRS = False
    Set QGRID_FASISL.InitializationClass = cInit
    With QGRID_FASISL
        .CustomDrawCellEnabled = True
        .DataFormatEnabled = True

'id, ArticoloPadre, Articolo, Descrizione, Qta, Spessore, Grammatura, Grammi_Metro, KGPezzi, PosizioneDa, PosizioneA, Colla
        '.INIT_ADDColumnSELECTIONBOXEXT "SELEZIONATO", "Sel.", gedTextEdit, 0, 0, , True, True
        .INIT_ADDColumn "IT01_SEQ", "Sequenza", gedTextEdit, 800, True
        .INIT_ADDColumn "IT01_CODFASE", "Fase", gedTextEdit, 1000, True
        .INIT_ADDColumn "IT01_DESCR", "Descrizione Fase", gedTextEdit, 3500, True
        .INIT_ADDColumn "IT01_CODREP_PD07", "Reparto", gedTextEdit, 750, True
        .INIT_ADDColumn "IT01_MACCHINA_PD08", "Macchina", gedTextEdit, 750, True
                               
        .InitializeSTART
                 .EnableEditing = True
                 .MODCOL_Editing "IT01_SEQ", Enable
        .InitializeEND

    End With
    

    
    
    
    Set cInit = Nothing
End Sub

Private Sub INITGRID_ARTICOLI()
    Dim strSQL As String
    Dim ors                As ADODB.Recordset
    
    Dim cInit As New TMS_QGRID.CLSFW_INITQGRID
    Set cInit.ActiveInterface = ActiveInterface
    Set cInit.RSColumns = Nothing
    cInit.KeyField = "Articolo"
    cInit.CreateColumnsFromRS = False
    Set QGRID_ARTICOLI.InitializationClass = cInit
    With QGRID_ARTICOLI
        .CustomDrawCellEnabled = True
        .DataFormatEnabled = True

'id, ArticoloPadre, Articolo, Descrizione, Qta, Spessore, Grammatura, Grammi_Metro, KGPezzi, PosizioneDa, PosizioneA, Colla
        '.INIT_ADDColumnSELECTIONBOXEXT "SELEZIONATO", "Sel.", gedTextEdit, 0, 0, , True, True
        .INIT_ADDColumn "Articolo", "Articolo", gedTextEdit, 2000, True
        .INIT_ADDColumn "Descrizione", "Descrizione", gedTextEdit, 3500, True
        .INIT_ADDColumn "PosizioneDa", "PosizioneDa", gedTextEdit, 1000, True
        .INIT_ADDColumn "PosizioneA", "PosizioneA", gedTextEdit, 1000, True
        .INIT_ADDColumn "Qta", "Qta", gedTextEdit, 800, True
        
        .INIT_ADDColumn "Spessore", "Spessore", gedTextEdit, 1000, True
        
        
        .INIT_ADDColumn "Grammatura", "Grammatura", gedTextEdit, 1000, True
        .INIT_ADDColumn "SommaSpessori", "SommaSpessori", gedTextEdit, 1000, True
        .INIT_ADDColumn "Grammi_Metro", "Grammi Metro", gedTextEdit, 1000, True
        .INIT_ADDColumn "KGPezzi", "KG/Pezzo", gedTextEdit, 1000, True
        .INIT_ADDColumn "Colla", "Colla", gedTextEdit, 1000, False
                
        .InitializeSTART
                 .EnableEditing = True
                 .MODCOL_Editing "PosizioneDa", Enable
                 .MODCOL_Editing "PosizioneA", Enable
                 .MODCOL_Editing "Qta", Enable
        .InitializeEND

    End With
    

    
    
    
    Set cInit = Nothing
End Sub








Private Sub CMD_LISTINI_Click()


Set Cls_ConnectMagazzino.ActiveInterface = ActiveInterface
        
        Cls_ConnectMagazzino.Left = 10
        Cls_ConnectMagazzino.Top = 1000
        

                                                                
        Call Cls_ConnectMagazzino.InterrogazioneListiniAttivi(RTrimN(TXT_CODART.Text), "")
        ActiveInterface.IsActive = True
        Set Cls_ConnectMagazzino.ActiveInterface = Nothing
        Set ActiveInterface.ClsGlobal.ActiveInterface = ActiveInterface

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
        TXT_CODART.SetTextFocus
    End If
    'Disattivo il messaggio a richiesta di aggiornare i dati modificati
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
    CMB_TIPOQTA.Enabled = True
    
  '  Call ReinizializzaVirtualFrame
    
    'Enzo 200703 Pulisci campi nuovi
    TXT_PZ.Text = ""
    
    TXT_DESGRUSTAT1.Text = ""
    
    
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
            Grst_RecSet_LI11_appendACQ.Fields("LI11_CODICE_CG08").Value = NVL(Grst_RecSet_LI11ACQ.Fields("DO11_VALUTA_CG08").Value, "")
            Grst_RecSet_LI11_appendACQ.Fields("LI11_DATACAMBIO").Value = NVL(Grst_RecSet_LI11ACQ.Fields("DO11_DATACAMBIO").Value, "")
            Grst_RecSet_LI11_appendACQ.Fields("LI11_CAMBIO").Value = NVL(Grst_RecSet_LI11ACQ.Fields("DO11_CAMBIO").Value, "")
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






Private Sub CMD_PARTITARI_Click()
Set Cls_ConnectMagazzino.ActiveInterface = ActiveInterface
        
        Cls_ConnectMagazzino.Left = 10
        Cls_ConnectMagazzino.Top = 1000
            Call Cls_ConnectMagazzino.InterrogazionePartitari(RTrimN(TXT_CODART.Text), "")
            
            

        ActiveInterface.IsActive = True
        Set Cls_ConnectMagazzino.ActiveInterface = Nothing
        Set ActiveInterface.ClsGlobal.ActiveInterface = ActiveInterface
End Sub

Private Sub CMD_RICARICA_Click()

End Sub

Private Sub CMD_SITGIAC_Click()
'
Set Cls_ConnectMagazzino.ActiveInterface = ActiveInterface
        
        Cls_ConnectMagazzino.Left = 10
        Cls_ConnectMagazzino.Top = 1000
        

                                                                
        Call Cls_ConnectMagazzino.InterrogazioneSituazioneGiacenze(RTrimN(TXT_CODART.Text))
        ActiveInterface.IsActive = True
        Set Cls_ConnectMagazzino.ActiveInterface = Nothing
        Set ActiveInterface.ClsGlobal.ActiveInterface = ActiveInterface

End Sub





Private Sub MDIActiveX1_FormLoad()
    On Error Resume Next
    MDIActiveX1.Move _
        IIf((ActiveInterface.Left = 0), 0, ActiveInterface.Left), _
        IIf((ActiveInterface.Top = 0), 0, ActiveInterface.Top), _
        21810, _
        13545
        
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
    
    
    
     

    
    If ActiveInterface.IsCalled Then
        TXT_CODART.Text = RTrimN(ActiveClass.CodiceArticolo)
        If RTrimN(ActiveClass.Opzione) <> "" Then
            DoEvents
        End If
        Old_Articolo = RTrimN(ActiveClass.CodiceArticolo)
        TXT_CODART.Enabled = False
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
    

    
    ' Istanzio la classe formattazione grid e data format per la ditta
    Set Pcls_GridFormat = New CLSFW_DataGrid
    
    ClickNuovo = False

    '
    ' Istanzio la classe formattazione grid e data format per la ditta
    '
    Set Pcls_GridFormat = New CLSFW_DataGrid
    Set Pstd_Format = New StdDataFormat
    Set Pstd_FormatDEP = New StdDataFormat
    
   
    
    OnUnload = False
    '
    ' Hasanin, 29/05/2006
    '
    IsLOaded = False
    
    Set TXT_CODART.ActiveInterface = ActiveInterface
    Set TXT_CODART.connessione = Gcon_Connect
    TXT_CODART.Ditta = Gstr_DittaCorrente

    Call TXT_CODART.MenuEntry("1", "Articoli movimentati", True)
    
    TMS_RESIZEFORM1.Initialize

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




Private Sub AggiornaTotali()

Dim strSQL As String

TXT_DESTERNO.Text = FormatNumber(GetValFromQuery("select (SUM(Spessore * Qta ) / 1000) * 2 as Desterno from IT00_CONFDISTINTE where Qta > 0"), 2)
'TXT_TOTKGCARTA.Text = GetValFromQuery("select SUM(Grammatura * Qta ) as KGPezzi from IT00_CONFDISTINTE where Qta > 0")
TXT_DMEDIO.Text = FormatNumber((CDbl(NVL(TXT_DESTERNO.Text, 0)) + CDbl(NVL(TXT_DINTERNO.Text, 0))) / 2, 2)
TXT_TOTCARTE.Text = FormatNumber(GetValFromQuery("select SUM(Qta ) as TotCarte from IT00_CONFDISTINTE where Qta > 0"), 2)

End Sub




Private Sub TMS_FLATBUTTON2_Click()
    
    Dim strSQL As String
    Dim TotGrammiCarta As Double
    
    TotGrammiCarta = CDbl(CDbl(NVL(TXT_TOTKGCARTA.Text, 0)) / 1000)
    ' Grammi_metro
    strSQL = " Update IT00_CONFDISTINTE  "
    strSQL = strSQL & "  "
    strSQL = strSQL & " set Grammi_Metro = (" & SQLDouble(NVL(TXT_DMEDIO.Text, 0)) & " * (3.14) ) * ( Grammatura / 1000 ) * isnull(qta,0) "
    strSQL = strSQL & "  "
    strSQL = strSQL & " from IT00_CONFDISTINTE  "
    
    
    Gcon_Connect.Execute strSQL
    
    
    'KG/Grammatura
        strSQL = " Update IT00_CONFDISTINTE  "
    strSQL = strSQL & "  "
    strSQL = strSQL & " set KGPezzi = (" & SQLDouble(NVL(TXT_LUNGHEZZA.Text, 0)) & " * Isnull(Grammi_metro,0) / 1000000 )"
    strSQL = strSQL & "  "
    strSQL = strSQL & " from IT00_CONFDISTINTE  "
    Gcon_Connect.Execute strSQL
    
    'Somma spessori
    strSQL = " Update IT00_CONFDISTINTE  "
    strSQL = strSQL & "  "
    strSQL = strSQL & " set SommaSpessori = ( Spessore * isnull(qta,0) ) / 1000"
    strSQL = strSQL & "  "
    strSQL = strSQL & " from IT00_CONFDISTINTE  "
    Gcon_Connect.Execute strSQL
    
    TXT_TOTSPESSORI.Text = GetValFromQuery("select SUM(SommaSpessori ) as SommaSpessori from IT00_CONFDISTINTE where Qta > 0")
    
    
    TXT_TOTKGCARTA.Text = FormatNumber(GetValFromQuery("select SUM(Grammi_Metro ) as SommaGrammi_Metro from IT00_CONFDISTINTE where Qta > 0") * CDbl(NVL(TXT_LUNGHEZZA.Text, 0)) / 1000000, 2)
    TXT_TOTCOLLA.Text = FormatNumber(CDbl(NVL(TXT_TOTKGCARTA.Text, 0)) * 12 / 100, 2)
    
    TXT_PESOTUBOU.Text = FormatNumber(CDbl(TXT_TOTKGCARTA.Text) + CDbl(TXT_TOTCOLLA.Text), 2)
    TXT_PESOTUBOA.Text = FormatNumber(CDbl(TXT_PESOTUBOU.Text) * 96 / 100, 2)
    
    TXT_PESOU.Text = FormatNumber((CDbl(TXT_PESOTUBOU.Text) / CDbl(TXT_LUNGHEZZA.Text)) * 1000, 2)
    TXT_PESOA.Text = FormatNumber(CDbl(TXT_PESOU.Text) * 96 / 100, 2)
    
    Call ImpostaVirtualFrame

   
End Sub

Private Sub TMS_QGRIDWRAPPER1_RowChanged()

End Sub

Private Sub TMS_FLATBUTTON5_Click()
Call CallDisba(TXT_CODART.Text, "", 0, "")
End Sub

Private Sub TMS_RESIZEFORM1_BeforeAutoInitialize(DisableAutoInitialize As Boolean)

    On Error Resume Next

    DisableAutoInitialize = True
    
    
    TMS_RESIZEFORM1.AddControl TMS_ANAGRAFICI, tsAnchorRight Or tsAnchorTop Or tsAnchorleft
    
    
    
    
    '
    
    
    '
    'TMS_RESIZEFORM1.AddControl TMS_SSTAB1, tsAnchorRight Or tsAnchorTop
    
    Err.Clear

End Sub


Public Function SQLDate(ByVal vdatData As Date) As String
  SQLDate = "CONVERT(DATETIME, '" & Year(vdatData) & "-" & _
                             Format(Month(vdatData), "00") & "-" & _
                             Format(Day(vdatData), "00") & _
                             " " & Format(Hour(vdatData), "00") & ":" & _
                             Format(Minute(vdatData), "00") & ":" & _
                             Format(Second(vdatData), "00") & "', 102)"
End Function

Public Function SQLDouble(ByVal vdblValue As Double) As String
  SQLDouble = Replace(vdblValue, ",", ".")
End Function

Public Function SQLString(ByVal vStr As String) As String
  SQLString = Replace(vStr, "'", "''")
End Function


Private Sub TXT_ARTICOLOPADREA_BeforeItem(Cancel As Boolean)

End Sub

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
     Case "Kgestione"
            Set Cls_ConnectMagazzino.ActiveInterface = ActiveInterface
            Cls_ConnectMagazzino.Left = 5
            Cls_ConnectMagazzino.Top = 1000
            Set Cls_ConnectMagazzino.ConnectField = Nothing
            Call Cls_ConnectMagazzino.ArticoloAnagrafica(RTrimN(TXT_CODART.Text))
            ActiveInterface.IsActive = True
            Set Cls_ConnectMagazzino.ActiveInterface = Nothing
            Set ActiveInterface.ClsGlobal.ActiveInterface = ActiveInterface
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






Public Sub Psub_Reinizializza()
    On Error GoTo Err:
    
    TXT_CODART.Text = ""
    TXT_DESCART.Text = ""
    TXT_DESCART.Default = ""
    
    'Enzo 200703 - Anagrafica estesa
    TXT_DESCARTEST.Text = ""
    TXT_DESCARTEST.Default = ""
    
    'Enzo 200703 - Carichi e scarichi
    
    
    CMB_TIPOQTA.Text = 0
    TXT_FAM.Text = ""
    TXT_SFAM.Text = ""
    TXT_GRUP.Text = ""
    TXT_SGRUP.Text = ""
    TXT_UM1.Text = ""
'    CHK_MOV.Text = 1
    
    

    
    Set QGRID_ARTICOLI.DataSource = Nothing
    
    QGRID_ARTICOLI.Refresh
    
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

Private Sub TxtEdit1_BeforeItem(Cancel As Boolean)

End Sub

Private Sub TXT_MACCHINA_FPF_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
  Cancel = False

  str_SQL = " SELECT PD08_MACCHINA, PD08_DESCR from PD08_MACCHINA "
      
  ReDim Arr_Fields(0 To 1, 0 To 1)
  Arr_Fields(0, 0) = "Codice"
  Arr_Fields(0, 1) = ""
  Arr_Fields(1, 0) = "Descrizione"
  Arr_Fields(1, 1) = ""
  
  Str_Caption = "Macchine Fasi Prodotto Finito"
  Str_Connect = Gstr_Connect
  TXT_MACCHINA_FPF.IDLookup = "lkp_Macchine"
End Sub

Private Sub TXT_MACCHINA_FSL_StartLookup(Cancel As Boolean, str_SQL As String, Arr_Fields As Variant, Str_Caption As String, Str_Connect As String)
 Cancel = False

  str_SQL = " SELECT PD08_MACCHINA, PD08_DESCR from PD08_MACCHINA "
      
  ReDim Arr_Fields(0 To 1, 0 To 1)
  Arr_Fields(0, 0) = "Codice"
  Arr_Fields(0, 1) = ""
  Arr_Fields(1, 0) = "Descrizione"
  Arr_Fields(1, 1) = ""
  
  Str_Caption = "Macchine Fasi Semi Lavorato"
  Str_Connect = Gstr_Connect
  TXT_MACCHINA_FSL.IDLookup = "lkp_MacchineSL"
End Sub
