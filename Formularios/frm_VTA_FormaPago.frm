VERSION 5.00
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form frm_VTA_FormaPago 
   BorderStyle     =   0  'None
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9825
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   9825
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMontoTotal 
      Enabled         =   0   'False
      Height          =   405
      Left            =   1800
      TabIndex        =   8
      Top             =   6720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtSecuencia 
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   6720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   495
      Left            =   5520
      TabIndex        =   6
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin TrueDBGrid70.TDBGrid GrdListaFP 
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   4683
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Codigo"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Forma Pago"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Codigo"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Pago"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Importe"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "CodTarj"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "CodMoneda"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "CodDocPago"
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "CodBanco"
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "CodDonacion"
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "CodCtaCte"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "TC"
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "NumTarj"
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "NumCuota"
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "FchVenc"
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "FlgCuotaNormal"
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "NumDocDesc"
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "NumMov"
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "FlgVaucherManual"
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "FechaMov"
      Columns(19).DataField=   ""
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "NumAutoriza"
      Columns(20).DataField=   ""
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "FchDoc"
      Columns(21).DataField=   ""
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "NumNC"
      Columns(22).DataField=   ""
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).Caption=   "NumCheque"
      Columns(23).DataField=   ""
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(24)._VlistStyle=   0
      Columns(24)._MaxComboItems=   5
      Columns(24).Caption=   "NomCli"
      Columns(24).DataField=   ""
      Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(25)._VlistStyle=   0
      Columns(25)._MaxComboItems=   5
      Columns(25).Caption=   "DniCli"
      Columns(25).DataField=   ""
      Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(26)._VlistStyle=   0
      Columns(26)._MaxComboItems=   5
      Columns(26).Caption=   "CodProd"
      Columns(26).DataField=   ""
      Columns(26)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(27)._VlistStyle=   0
      Columns(27)._MaxComboItems=   5
      Columns(27).Caption=   "CodBtl"
      Columns(27).DataField=   ""
      Columns(27)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(28)._VlistStyle=   0
      Columns(28)._MaxComboItems=   5
      Columns(28).Caption=   "DocRef"
      Columns(28).DataField=   ""
      Columns(28)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(29)._VlistStyle=   0
      Columns(29)._MaxComboItems=   5
      Columns(29).Caption=   "RetEfect"
      Columns(29).DataField=   ""
      Columns(29)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(30)._VlistStyle=   0
      Columns(30)._MaxComboItems=   5
      Columns(30).Caption=   "ObsCheque"
      Columns(30).DataField=   ""
      Columns(30)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(31)._VlistStyle=   0
      Columns(31)._MaxComboItems=   5
      Columns(31).Caption=   "FlgRedondeo"
      Columns(31).DataField=   ""
      Columns(31)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(32)._VlistStyle=   0
      Columns(32)._MaxComboItems=   5
      Columns(32).Caption=   "ImporteDol"
      Columns(32).DataField=   "ImporteDol"
      Columns(32)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(33)._VlistStyle=   0
      Columns(33)._MaxComboItems=   5
      Columns(33).Caption=   "NombTitular"
      Columns(33).DataField=   "NombTitular"
      Columns(33)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(34)._VlistStyle=   0
      Columns(34)._MaxComboItems=   5
      Columns(34).Caption=   "Número Vale"
      Columns(34).DataField=   "NumVale"
      Columns(34)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   35
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=35"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1270"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1191"
      Splits(0)._ColumnProps(4)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=4498"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4419"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=185"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=106"
      Splits(0)._ColumnProps(13)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(3).Width=3757"
      Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=3678"
      Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(19)=   "Column(4).Width=1799"
      Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=1720"
      Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(23)=   "Column(5).Width=1191"
      Splits(0)._ColumnProps(24)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(5)._WidthInPix=1111"
      Splits(0)._ColumnProps(26)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(27)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(28)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(29)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(31)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(32)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(33)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(34)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(36)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(37)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(38)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(39)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(41)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(42)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(43)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(44)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(46)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(47)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(48)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(49)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(50)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(51)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(52)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(53)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(54)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(56)=   "Column(11).Visible=0"
      Splits(0)._ColumnProps(57)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(58)=   "Column(12).Width=2725"
      Splits(0)._ColumnProps(59)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(60)=   "Column(12)._WidthInPix=2646"
      Splits(0)._ColumnProps(61)=   "Column(12).Visible=0"
      Splits(0)._ColumnProps(62)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(63)=   "Column(13).Width=2725"
      Splits(0)._ColumnProps(64)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(65)=   "Column(13)._WidthInPix=2646"
      Splits(0)._ColumnProps(66)=   "Column(13).Visible=0"
      Splits(0)._ColumnProps(67)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(68)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(69)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(70)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(71)=   "Column(14).Visible=0"
      Splits(0)._ColumnProps(72)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(73)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(74)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(75)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(76)=   "Column(15).Visible=0"
      Splits(0)._ColumnProps(77)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(78)=   "Column(16).Width=5292"
      Splits(0)._ColumnProps(79)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(80)=   "Column(16)._WidthInPix=5212"
      Splits(0)._ColumnProps(81)=   "Column(16).Visible=0"
      Splits(0)._ColumnProps(82)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(83)=   "Column(17).Width=2725"
      Splits(0)._ColumnProps(84)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(85)=   "Column(17)._WidthInPix=2646"
      Splits(0)._ColumnProps(86)=   "Column(17).Visible=0"
      Splits(0)._ColumnProps(87)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(88)=   "Column(18).Width=2725"
      Splits(0)._ColumnProps(89)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(90)=   "Column(18)._WidthInPix=2646"
      Splits(0)._ColumnProps(91)=   "Column(18).Visible=0"
      Splits(0)._ColumnProps(92)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(93)=   "Column(19).Width=2725"
      Splits(0)._ColumnProps(94)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(95)=   "Column(19)._WidthInPix=2646"
      Splits(0)._ColumnProps(96)=   "Column(19).Visible=0"
      Splits(0)._ColumnProps(97)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(98)=   "Column(20).Width=2725"
      Splits(0)._ColumnProps(99)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(100)=   "Column(20)._WidthInPix=2646"
      Splits(0)._ColumnProps(101)=   "Column(20).Visible=0"
      Splits(0)._ColumnProps(102)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(103)=   "Column(21).Width=2725"
      Splits(0)._ColumnProps(104)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(105)=   "Column(21)._WidthInPix=2646"
      Splits(0)._ColumnProps(106)=   "Column(21).Visible=0"
      Splits(0)._ColumnProps(107)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(108)=   "Column(22).Width=2725"
      Splits(0)._ColumnProps(109)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(110)=   "Column(22)._WidthInPix=2646"
      Splits(0)._ColumnProps(111)=   "Column(22).Visible=0"
      Splits(0)._ColumnProps(112)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(113)=   "Column(23).Width=2725"
      Splits(0)._ColumnProps(114)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(115)=   "Column(23)._WidthInPix=2646"
      Splits(0)._ColumnProps(116)=   "Column(23).Visible=0"
      Splits(0)._ColumnProps(117)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(118)=   "Column(24).Width=2725"
      Splits(0)._ColumnProps(119)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(120)=   "Column(24)._WidthInPix=2646"
      Splits(0)._ColumnProps(121)=   "Column(24).Visible=0"
      Splits(0)._ColumnProps(122)=   "Column(24).Order=25"
      Splits(0)._ColumnProps(123)=   "Column(25).Width=2725"
      Splits(0)._ColumnProps(124)=   "Column(25).DividerColor=0"
      Splits(0)._ColumnProps(125)=   "Column(25)._WidthInPix=2646"
      Splits(0)._ColumnProps(126)=   "Column(25).Visible=0"
      Splits(0)._ColumnProps(127)=   "Column(25).Order=26"
      Splits(0)._ColumnProps(128)=   "Column(26).Width=2725"
      Splits(0)._ColumnProps(129)=   "Column(26).DividerColor=0"
      Splits(0)._ColumnProps(130)=   "Column(26)._WidthInPix=2646"
      Splits(0)._ColumnProps(131)=   "Column(26).Visible=0"
      Splits(0)._ColumnProps(132)=   "Column(26).Order=27"
      Splits(0)._ColumnProps(133)=   "Column(27).Width=2725"
      Splits(0)._ColumnProps(134)=   "Column(27).DividerColor=0"
      Splits(0)._ColumnProps(135)=   "Column(27)._WidthInPix=2646"
      Splits(0)._ColumnProps(136)=   "Column(27).Visible=0"
      Splits(0)._ColumnProps(137)=   "Column(27).Order=28"
      Splits(0)._ColumnProps(138)=   "Column(28).Width=2725"
      Splits(0)._ColumnProps(139)=   "Column(28).DividerColor=0"
      Splits(0)._ColumnProps(140)=   "Column(28)._WidthInPix=2646"
      Splits(0)._ColumnProps(141)=   "Column(28).Visible=0"
      Splits(0)._ColumnProps(142)=   "Column(28).Order=29"
      Splits(0)._ColumnProps(143)=   "Column(29).Width=2725"
      Splits(0)._ColumnProps(144)=   "Column(29).DividerColor=0"
      Splits(0)._ColumnProps(145)=   "Column(29)._WidthInPix=2646"
      Splits(0)._ColumnProps(146)=   "Column(29).Visible=0"
      Splits(0)._ColumnProps(147)=   "Column(29).Order=30"
      Splits(0)._ColumnProps(148)=   "Column(30).Width=2725"
      Splits(0)._ColumnProps(149)=   "Column(30).DividerColor=0"
      Splits(0)._ColumnProps(150)=   "Column(30)._WidthInPix=2646"
      Splits(0)._ColumnProps(151)=   "Column(30).Visible=0"
      Splits(0)._ColumnProps(152)=   "Column(30).Order=31"
      Splits(0)._ColumnProps(153)=   "Column(31).Width=2725"
      Splits(0)._ColumnProps(154)=   "Column(31).DividerColor=0"
      Splits(0)._ColumnProps(155)=   "Column(31)._WidthInPix=2646"
      Splits(0)._ColumnProps(156)=   "Column(31).Visible=0"
      Splits(0)._ColumnProps(157)=   "Column(31).Order=32"
      Splits(0)._ColumnProps(158)=   "Column(32).Width=5292"
      Splits(0)._ColumnProps(159)=   "Column(32).DividerColor=0"
      Splits(0)._ColumnProps(160)=   "Column(32)._WidthInPix=5212"
      Splits(0)._ColumnProps(161)=   "Column(32).Visible=0"
      Splits(0)._ColumnProps(162)=   "Column(32).Order=33"
      Splits(0)._ColumnProps(163)=   "Column(33).Width=2725"
      Splits(0)._ColumnProps(164)=   "Column(33).DividerColor=0"
      Splits(0)._ColumnProps(165)=   "Column(33)._WidthInPix=2646"
      Splits(0)._ColumnProps(166)=   "Column(33).Visible=0"
      Splits(0)._ColumnProps(167)=   "Column(33).Order=34"
      Splits(0)._ColumnProps(168)=   "Column(34).Width=6694"
      Splits(0)._ColumnProps(169)=   "Column(34).DividerColor=0"
      Splits(0)._ColumnProps(170)=   "Column(34)._WidthInPix=6615"
      Splits(0)._ColumnProps(171)=   "Column(34).Order=35"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      Appearance      =   2
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   13160660
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=111,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=78,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=86,.parent=13"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=90,.parent=13"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=87,.parent=14"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=88,.parent=15"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=89,.parent=17"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=94,.parent=13"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=91,.parent=14"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=92,.parent=15"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=93,.parent=17"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=98,.parent=13"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=95,.parent=14"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=96,.parent=15"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=97,.parent=17"
      _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=102,.parent=13"
      _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=99,.parent=14"
      _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=100,.parent=15"
      _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=101,.parent=17"
      _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=106,.parent=13"
      _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=103,.parent=14"
      _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=104,.parent=15"
      _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=105,.parent=17"
      _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=114,.parent=13"
      _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=111,.parent=14"
      _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=112,.parent=15"
      _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=113,.parent=17"
      _StyleDefs(112) =   "Splits(0).Columns(19).Style:id=118,.parent=13"
      _StyleDefs(113) =   "Splits(0).Columns(19).HeadingStyle:id=115,.parent=14"
      _StyleDefs(114) =   "Splits(0).Columns(19).FooterStyle:id=116,.parent=15"
      _StyleDefs(115) =   "Splits(0).Columns(19).EditorStyle:id=117,.parent=17"
      _StyleDefs(116) =   "Splits(0).Columns(20).Style:id=122,.parent=13"
      _StyleDefs(117) =   "Splits(0).Columns(20).HeadingStyle:id=119,.parent=14"
      _StyleDefs(118) =   "Splits(0).Columns(20).FooterStyle:id=120,.parent=15"
      _StyleDefs(119) =   "Splits(0).Columns(20).EditorStyle:id=121,.parent=17"
      _StyleDefs(120) =   "Splits(0).Columns(21).Style:id=126,.parent=13"
      _StyleDefs(121) =   "Splits(0).Columns(21).HeadingStyle:id=123,.parent=14"
      _StyleDefs(122) =   "Splits(0).Columns(21).FooterStyle:id=124,.parent=15"
      _StyleDefs(123) =   "Splits(0).Columns(21).EditorStyle:id=125,.parent=17"
      _StyleDefs(124) =   "Splits(0).Columns(22).Style:id=130,.parent=13"
      _StyleDefs(125) =   "Splits(0).Columns(22).HeadingStyle:id=127,.parent=14"
      _StyleDefs(126) =   "Splits(0).Columns(22).FooterStyle:id=128,.parent=15"
      _StyleDefs(127) =   "Splits(0).Columns(22).EditorStyle:id=129,.parent=17"
      _StyleDefs(128) =   "Splits(0).Columns(23).Style:id=134,.parent=13"
      _StyleDefs(129) =   "Splits(0).Columns(23).HeadingStyle:id=131,.parent=14"
      _StyleDefs(130) =   "Splits(0).Columns(23).FooterStyle:id=132,.parent=15"
      _StyleDefs(131) =   "Splits(0).Columns(23).EditorStyle:id=133,.parent=17"
      _StyleDefs(132) =   "Splits(0).Columns(24).Style:id=138,.parent=13"
      _StyleDefs(133) =   "Splits(0).Columns(24).HeadingStyle:id=135,.parent=14"
      _StyleDefs(134) =   "Splits(0).Columns(24).FooterStyle:id=136,.parent=15"
      _StyleDefs(135) =   "Splits(0).Columns(24).EditorStyle:id=137,.parent=17"
      _StyleDefs(136) =   "Splits(0).Columns(25).Style:id=142,.parent=13"
      _StyleDefs(137) =   "Splits(0).Columns(25).HeadingStyle:id=139,.parent=14"
      _StyleDefs(138) =   "Splits(0).Columns(25).FooterStyle:id=140,.parent=15"
      _StyleDefs(139) =   "Splits(0).Columns(25).EditorStyle:id=141,.parent=17"
      _StyleDefs(140) =   "Splits(0).Columns(26).Style:id=146,.parent=13"
      _StyleDefs(141) =   "Splits(0).Columns(26).HeadingStyle:id=143,.parent=14"
      _StyleDefs(142) =   "Splits(0).Columns(26).FooterStyle:id=144,.parent=15"
      _StyleDefs(143) =   "Splits(0).Columns(26).EditorStyle:id=145,.parent=17"
      _StyleDefs(144) =   "Splits(0).Columns(27).Style:id=110,.parent=13"
      _StyleDefs(145) =   "Splits(0).Columns(27).HeadingStyle:id=107,.parent=14"
      _StyleDefs(146) =   "Splits(0).Columns(27).FooterStyle:id=108,.parent=15"
      _StyleDefs(147) =   "Splits(0).Columns(27).EditorStyle:id=109,.parent=17"
      _StyleDefs(148) =   "Splits(0).Columns(28).Style:id=150,.parent=13"
      _StyleDefs(149) =   "Splits(0).Columns(28).HeadingStyle:id=147,.parent=14"
      _StyleDefs(150) =   "Splits(0).Columns(28).FooterStyle:id=148,.parent=15"
      _StyleDefs(151) =   "Splits(0).Columns(28).EditorStyle:id=149,.parent=17"
      _StyleDefs(152) =   "Splits(0).Columns(29).Style:id=158,.parent=13"
      _StyleDefs(153) =   "Splits(0).Columns(29).HeadingStyle:id=155,.parent=14"
      _StyleDefs(154) =   "Splits(0).Columns(29).FooterStyle:id=156,.parent=15"
      _StyleDefs(155) =   "Splits(0).Columns(29).EditorStyle:id=157,.parent=17"
      _StyleDefs(156) =   "Splits(0).Columns(30).Style:id=162,.parent=13"
      _StyleDefs(157) =   "Splits(0).Columns(30).HeadingStyle:id=159,.parent=14"
      _StyleDefs(158) =   "Splits(0).Columns(30).FooterStyle:id=160,.parent=15"
      _StyleDefs(159) =   "Splits(0).Columns(30).EditorStyle:id=161,.parent=17"
      _StyleDefs(160) =   "Splits(0).Columns(31).Style:id=154,.parent=13"
      _StyleDefs(161) =   "Splits(0).Columns(31).HeadingStyle:id=151,.parent=14"
      _StyleDefs(162) =   "Splits(0).Columns(31).FooterStyle:id=152,.parent=15"
      _StyleDefs(163) =   "Splits(0).Columns(31).EditorStyle:id=153,.parent=17"
      _StyleDefs(164) =   "Splits(0).Columns(32).Style:id=166,.parent=13"
      _StyleDefs(165) =   "Splits(0).Columns(32).HeadingStyle:id=163,.parent=14"
      _StyleDefs(166) =   "Splits(0).Columns(32).FooterStyle:id=164,.parent=15"
      _StyleDefs(167) =   "Splits(0).Columns(32).EditorStyle:id=165,.parent=17"
      _StyleDefs(168) =   "Splits(0).Columns(33).Style:id=170,.parent=13"
      _StyleDefs(169) =   "Splits(0).Columns(33).HeadingStyle:id=167,.parent=14"
      _StyleDefs(170) =   "Splits(0).Columns(33).FooterStyle:id=168,.parent=15"
      _StyleDefs(171) =   "Splits(0).Columns(33).EditorStyle:id=169,.parent=17"
      _StyleDefs(172) =   "Splits(0).Columns(34).Style:id=174,.parent=13"
      _StyleDefs(173) =   "Splits(0).Columns(34).HeadingStyle:id=171,.parent=14"
      _StyleDefs(174) =   "Splits(0).Columns(34).FooterStyle:id=172,.parent=15"
      _StyleDefs(175) =   "Splits(0).Columns(34).EditorStyle:id=173,.parent=17"
      _StyleDefs(176) =   "Named:id=33:Normal"
      _StyleDefs(177) =   ":id=33,.parent=0"
      _StyleDefs(178) =   "Named:id=34:Heading"
      _StyleDefs(179) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(180) =   ":id=34,.wraptext=-1"
      _StyleDefs(181) =   "Named:id=35:Footing"
      _StyleDefs(182) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(183) =   "Named:id=36:Selected"
      _StyleDefs(184) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(185) =   "Named:id=37:Caption"
      _StyleDefs(186) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(187) =   "Named:id=38:HighlightRow"
      _StyleDefs(188) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(189) =   "Named:id=39:EvenRow"
      _StyleDefs(190) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(191) =   "Named:id=40:OddRow"
      _StyleDefs(192) =   ":id=40,.parent=33"
      _StyleDefs(193) =   "Named:id=41:RecordSelector"
      _StyleDefs(194) =   ":id=41,.parent=34"
      _StyleDefs(195) =   "Named:id=42:FilterBar"
      _StyleDefs(196) =   ":id=42,.parent=33"
      _StyleDefs(197) =   "Named:id=0:"
      _StyleDefs(198) =   ":id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(199) =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(200) =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(201) =   ":id=0,.borderColor=&H80000005&,.borderType=111,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(202) =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(203) =   ":id=0,.fontname=MS Sans Serif"
   End
   Begin vbp_Ventas.ctlGrilla grdFormaPago 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9615
      _extentx        =   16960
      _extenty        =   4895
      menupopup       =   0   'False
      resalte         =   0   'False
   End
   Begin VB.Label lblDocumento 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   3660
      Width           =   180
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "F1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   420
      Width           =   180
   End
   Begin VB.Label Label2 
      Caption         =   "Formas de pago registradas : "
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   -240
      X2              =   9720
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line2 
      X1              =   -180
      X2              =   9720
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Forma de Pago"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Index           =   4
      Left            =   420
      TabIndex        =   1
      Top             =   60
      Width           =   1590
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   60
      Picture         =   "frm_VTA_FormaPago.frx":0000
      Top             =   60
      Width           =   240
   End
End
Attribute VB_Name = "frm_VTA_FormaPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Param_Tipo_Documento As String
Public Param_Numero_Documento As String
Public Modificacion As Boolean
Dim objTarjeta As New clsFormaPago
Dim objFormaPago As New clsFormaPago
Dim strFormaPago As String
'Public pstrDato As String
'Public pstrDatoDes As String
'Public pstrFPago As String
Public cCodFPadre As String
Public cCodFHijo As String
Dim dblPagar As Double

Private Sub cmdAceptar_Click()
On Error GoTo Control
    
    

Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number

End Sub

Private Sub cmdGrabar_Click()
On Error GoTo handle
Dim objFormaPago As New clsFormaPago
Dim strMensaje As String
    If objVenta.FormaPago.UpperBound(1) < 0 Then MsgBox "No se puede grabar sin ninguna forma de pago", vbCritical, App.ProductName: Exit Sub
    strMensaje = objFormaPago.Graba(Param_Tipo_Documento, Param_Numero_Documento, txtSecuencia.Text, txtMontoTotal.Text, objUsuario.Codigo)
    If strMensaje = "" Then
        MsgBox "Se actualizo satisfatoriamente la forma de pago", vbExclamation, App.ProductName
    End If
    Modificacion = False
    frm_VTA_ConsultaDoc.pblnFpago = False
   Set objFormaPago = Nothing
    objVenta.FormaPago.ReDim 0, -1, 0, 32
Unload Me
   frm_VTA_ConsultaDoc.SetFocus
   
  Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub Form_Activate()
    GrdListaFP.Array = objVenta.FormaPago
    GrdListaFP.Rebind
    'I.ECASTILLO 21.04.2021
    Dim flgModifica As String
    If Modificacion = True Then
        flgModifica = "1"
    Else
        flgModifica = ""
    End If
    Set grdFormaPago.DataSource = objFormaPago.ListaFPagoTipMaquina(objUsuario.TipoMaquina, flgModifica)
    Call SeteaGrilla
    'F.ECASTILLO 21.04.2021
'    If objVenta.FormaPago.UpperBound(1) <> "-1" Then
'        dblPagar = Val(frmPedido.lblTotalPagar.Caption)
'        frmPedido.lblPagado.Caption = Format(objVenta.TotalFormaPago, "#,###,##0.00")
'        frmPedido.lblVuelto.Caption = Format(Val(objVenta.TotalFormaPago - dblPagar), "#,###,##0.00")
'    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    psub_KeyDownAplicacion KeyCode, Shift
    
    Select Case KeyCode
        Case vbKeyF1
            grdFormaPago.SetFocus
        Case vbKeyF2
            GrdListaFP.SetFocus
        Case vbKeyEscape
            If Modificacion = True Then
                objVenta.FormaPago.ReDim 0, -1, 0, 33
            End If
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
Dim flgModifica As String
On Error GoTo Control

If Modificacion = True Then
   objVenta.FormaPago.ReDim 0, -1, 0, 33
    cmdGrabar.Visible = True
    lblDocumento.Visible = True
    txtSecuencia.Visible = True
    txtMontoTotal.Visible = True
    lblDocumento.Caption = " " & Param_Tipo_Documento & "-->N°" & Param_Numero_Documento
    Dim objFormaPago As New clsFormaPago
    Dim rs As oraDynaset
    Set rs = objFormaPago.ListaFormaPDocumento(Param_Tipo_Documento, Param_Numero_Documento)
    txtSecuencia.Text = "" & rs("SEC_FORPAG_DOC").Value
    txtMontoTotal.Text = "" & objFormaPago.ListaTotalSec(txtSecuencia.Text)("TOTAL").Value
    'txtMontoTotal.Text = "" & rs("MTO_TOTAL").Value
    While Not rs.EOF
    
        If rs("FLG_RETIRO_EFEC").Value <> "1" Then
        
            objVenta.AgregaFormaPago "" & rs("COD_FORMA_PAGO"), _
                                     "" & rs("DES_FORMA_PAGO"), _
                                     "" & rs("COD_HIJO"), _
                                     "" & rs("DES_HIJO"), _
                                     "" & rs("IMP_SIN_REDONDEO"), _
                                     "" & rs("COD_TIPO_TARJETA"), _
                                     "" & rs("COD_MONEDA"), _
                                     "" & rs("COD_DOCUMENTO_PAGO"), _
                                     "" & rs("COD_BANCO"), _
                                     "" & rs("COD_DONACION"), _
                                     "" & rs("COD_CTACTE_BTL"), _
                                     "" & rs("IMP_TIPO_CAMBIO"), _
                                     "" & rs("NUM_TARJETA"), _
                                     "" & rs("NUM_CUOTAS"), _
                                     "" & rs("FCH_VENCIMIENTO"), _
                                     "" & rs("FLG_CUOTA_NORMAL"), _
                                     "" & rs("NUM_DOCUMENTO_PAGO"), _
                                     "" & rs("NUM_MOVIMIENTO"), _
                                     "" & rs("FLG_VOUCHER_MANUAL"), _
                                     "" & rs("FCH_MOVIMIENTO"), _
                                     "" & rs("NUM_AUTORIZACION"), _
                                     "" & rs("FCH_DOC_NOTA_CRED"), _
                                     "" & rs("NUM_DOC_NOTA_CRED"), _
                                     "" & rs("NUM_CHEQUE"), "", "", "", "", "", _
                                     "" & rs("FLG_RETIRO_EFEC"), "", "", "", ""
        End If
    rs.MoveNext
    Wend
    Set objFormaPago = Nothing
    flgModifica = "1"
Else
    flgModifica = ""
    lblDocumento.Visible = False
    txtSecuencia.Visible = False
    txtMontoTotal.Visible = False
    cmdGrabar.Visible = False
End If
    Me.top = 0
    Me.left = 0
    setteaFormulario Me
    'pstrFPago = "1"
    'Set grdFormaPago.DataSource = ObjFormaPago.Lista
    'Dim objGetTimer As New cGetTimer
    
    'objGetTimer.StartTimer
    'Set grdFormaPago.DataSource = objFormaPago.ListaFPagoTipMaquina(objUsuario.TipoMaquina, flgModifica) 'ECASTILLO 21.04.2021
    'objGetTimer.StopTimer
    'Debug.Print objGetTimer.ElapsedTime
    'Call SeteaGrilla 'ECASTILLO 21.04.2021
    GrdListaFP.Columns(0).Visible = False
    GrdListaFP.Columns(2).Visible = False
    ''''''''GrdListaFP.Columns(12).Visible = True
    GrdListaFP.AllowUpdate = False
    GrdListaFP.MarqueeStyle = dbgHighlightRow

   
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error : " & Err.Number
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Modificacion = False
End Sub

Private Sub grdFormaPago_ButtonClick(ByVal ColIndex As Integer)
    Select Case ColIndex
        Case 1
            grdFormaPago_DblClick
    End Select
End Sub

Private Sub grdFormaPago_DblClick()
On Error GoTo handle
    'MsgBox "ECASTILLO 24.04.2020 - grdFormaPago_DblClick"
    frm_VTA_FormaPagoEfectivo.pblnOpc = False
    'MsgBox "ECASTILLO 24.04.2020 - Select Case"
    Select Case Trim(strFormaPago)
        Case "001"  'Efectivo
               frm_VTA_FormaPagoEfectivo.pstrDato = grdFormaPago.Columns("COD_FORMA_PAGO").Value
               frm_VTA_FormaPagoEfectivo.pstrDatoDes = grdFormaPago.Columns("DES_FORMA_PAGO").Value
               frm_VTA_FormaPagoEfectivo.Show
               frm_VTA_FormaPagoEfectivo.SetFocus
        Case "002"  'Tarjeta
               frm_VTA_FormaPagoTarjeta.pstrDato = grdFormaPago.Columns("COD_FORMA_PAGO").Value
               frm_VTA_FormaPagoTarjeta.pstrDatoDes = grdFormaPago.Columns("DES_FORMA_PAGO").Value
               frm_VTA_FormaPagoTarjeta.Show
               frm_VTA_FormaPagoTarjeta.SetFocus
        Case "003"  'Tarjeta
               frm_VTA_FormaPagoCredito.pstrDato = grdFormaPago.Columns("COD_FORMA_PAGO").Value
               frm_VTA_FormaPagoCredito.pstrDatoDes = grdFormaPago.Columns("DES_FORMA_PAGO").Value
               frm_VTA_FormaPagoCredito.Show
               frm_VTA_FormaPagoCredito.SetFocus
               
        Case "004"  ' NC
               frm_VTA_FormaPagoNC.pstrDato = grdFormaPago.Columns("COD_FORMA_PAGO").Value
               frm_VTA_FormaPagoNC.pstrDatoDes = grdFormaPago.Columns("DES_FORMA_PAGO").Value
               frm_VTA_FormaPagoNC.Show
               frm_VTA_FormaPagoNC.SetFocus
        Case "005"    'Deposito en Cuenta
               frm_VTA_FormaPagoDepCta.pstrDato = grdFormaPago.Columns("COD_FORMA_PAGO").Value
               frm_VTA_FormaPagoDepCta.pstrDatoDes = grdFormaPago.Columns("DES_FORMA_PAGO").Value
               frm_VTA_FormaPagoDepCta.Show
               frm_VTA_FormaPagoDepCta.SetFocus
        Case "006"  'Cheque
               frm_VTA_FormaPagoCheque.pstrDato = grdFormaPago.Columns("COD_FORMA_PAGO").Value
               frm_VTA_FormaPagoCheque.pstrDatoDes = grdFormaPago.Columns("DES_FORMA_PAGO").Value
               frm_VTA_FormaPagoCheque.Show
               frm_VTA_FormaPagoCheque.SetFocus
        Case "007"  'Documento Dcto
               frm_VTA_FormaPagoDD.pstrDato = grdFormaPago.Columns("COD_FORMA_PAGO").Value
               frm_VTA_FormaPagoDD.pstrDatoDes = grdFormaPago.Columns("DES_FORMA_PAGO").Value
               frm_VTA_FormaPagoDD.Show
               frm_VTA_FormaPagoDD.SetFocus
        Case "009"  'Donacion
               frm_VTA_FormaPagoDonacion.pstrDato = grdFormaPago.Columns("COD_FORMA_PAGO").Value
               frm_VTA_FormaPagoDonacion.pstrDatoDes = grdFormaPago.Columns("DES_FORMA_PAGO").Value
               frm_VTA_FormaPagoDonacion.Show
               frm_VTA_FormaPagoDonacion.SetFocus
        Case "010"  'Puntos
              'frm_VTA_FormaPagoDonacion.pstrDato = grdFormaPago.Columns("COD_FORMA_PAGO").Value
               'frm_VTA_FormaPagoDonacion.pstrDatoDes = grdFormaPago.Columns("DES_FORMA_PAGO").Value
               frm_VTA_Puntos.Show
               frm_VTA_Puntos.SetFocus
        Case "011"  'Vale Fid
               'MsgBox "ECASTILLO 24.04.2020 - Case Vale Fid"
               If Len(frmPedido.pstrCodCliente_Ink) = 0 Then MsgBox "Solo para clientes Afiliados.", vbCritical, App.ProductName: Exit Sub
               frm_VTA_FormaPagoVF.pstrDato = grdFormaPago.Columns("COD_FORMA_PAGO").Value
               frm_VTA_FormaPagoVF.pstrDatoDes = grdFormaPago.Columns("DES_FORMA_PAGO").Value
               'MsgBox "ECASTILLO 24.04.2020 - FrmVF.Show || " & frm_VTA_FormaPagoVF.pstrDato & " || " & frm_VTA_FormaPagoVF.pstrDatoDes
               frm_VTA_FormaPagoVF.Show
               'MsgBox "ECASTILLO 24.04.2020 - FrmVF.SetFocus"
               frm_VTA_FormaPagoVF.SetFocus
    End Select
    'MsgBox "ECASTILLO 24.04.2020 - End Select"
    Exit Sub
handle:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub

Private Sub grdFormaPago_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
                 grdFormaPago_DblClick
        Case vbKeySpace
                 grdFormaPago_DblClick
    End Select
End Sub

Private Sub grdFormaPago_RegistroSeleccionado(ByVal DatoColumna0 As String)
    strFormaPago = DatoColumna0
End Sub

Private Sub SeteaGrilla()
  Dim arrCampos As Variant
  Dim arrCaption As Variant
  Dim arrAncho As Variant
  Dim arrAlineacion As Variant
  
    arrCampos = Array("COD_FORMA_PAGO", "DES_FORMA_PAGO")
    arrCaption = Array("Codigo", "Descripción")
    arrAncho = Array(900, 4500)
    arrAlineacion = Array(dbgCenter, dbgLeft)
    grdFormaPago.FormatoGrilla arrCampos, arrCaption, arrAncho, arrAlineacion
    
    Dim i%
    For i = 0 To grdFormaPago.Columns.Count - 1
        grdFormaPago.Columns(i).Visible = False
    Next i
    grdFormaPago.Columns("COD_FORMA_PAGO").Visible = False
    grdFormaPago.Columns("DES_FORMA_PAGO").Visible = True
    
    grdFormaPago.Columns("DES_FORMA_PAGO").ButtonText = True
    grdFormaPago.Styles(5).ForeColor = vbBlack
    grdFormaPago.Styles(5).Font.Bold = True
    'grdFormaPago.Styles(5).BackColor = vbYellow
    
    grdFormaPago.Columns("COD_FORMA_PAGO").AllowFocus = False
    'grdFormaPago.RowHeight = 1.5 * grdFormaPago.RowHeight
End Sub


Private Sub GrdListaFP_DblClick()
    On Error GoTo CtrlErr
    ModificarFP
    Exit Sub
CtrlErr:
    MsgBox Err.Description, vbCritical, App.ProductName
End Sub


Private Sub GrdListaFP_KeyDown(KeyCode As Integer, Shift As Integer)
    
    frm_VTA_FormaPagoEfectivo.pblnOpc = True
    Select Case KeyCode
        Case vbKeyReturn
                ModificarFP
         Case vbKeyDelete
           On Error GoTo CtrlErr
            If GrdListaFP.Columns("FlgRedondeo").Value = "1" Then Exit Sub
            
            Dim strCodFormaPago As String
            Dim strCodFormaPagoHijo As String
            Dim StrNumTarjeta As String
            
            strCodFormaPago = GrdListaFP.Columns(0)
            strCodFormaPagoHijo = GrdListaFP.Columns(2)
            StrNumTarjeta = GrdListaFP.Columns(12)
            
            If strCodFormaPago = "002" And strCodFormaPagoHijo = "032" Then
               gintFidelizado = 0
            End If
            
            
            GrdListaFP.Delete
            
            
            'objVenta.RemoverFormaPago strCodFormaPago, strCodFormaPagoHijo, strNumTarjeta
            
            frmPedido.Cal_Promo
            frmPedido.Cal_Montos
            
            
            
            
            'GrdListaFP.Delete
            'dblPagar = Val(frmPedido.lblTotalPagar.Caption)
            'frmPedido.lblPagado.Caption = Format(objVenta.TotalFormaPago, "#,###,##0.00")
            'frmPedido.lblVuelto.Caption = Format(Val(objVenta.TotalFormaPago - dblPagar), "#,###,##0.00")
CtrlErr:
            On Error GoTo 0
    End Select
End Sub
Sub ModificarFP()
    Dim cImporte As Double
    Dim cTC As Double

            On Error GoTo CtlrErr
            If GrdListaFP.ApproxCount = 0 Then Exit Sub

            'Toma el codigo de forma de pago a cambiar'
            
            cCodFPadre = objVenta.FormaPago(GrdListaFP.row, 0)
            cCodFHijo = objVenta.FormaPago(GrdListaFP.row, 2)
            cImporte = objVenta.FormaPago(GrdListaFP.row, 4)
            cTC = objVenta.FormaPago(GrdListaFP.row, 11)
            'objVenta.dc_cod_forma_pago = "" & gclsOracle.FN_Valor("DELIVERY.PKG_DC_CAPPA.FN_GET_CODIGO_DIG", cCodFPadre, cCodFHijo) 'ECASTILLO 27.10.2020
            If cCodFPadre = "001" Then 'EFECTIVO'
                        frm_VTA_FormaPagoEfectivo.pstrDato = GrdListaFP.Columns(0).Value 'objVenta.FormaPago(GrdListaFP.Row, 0)
                        frm_VTA_FormaPagoEfectivo.pstrDatoDes = GrdListaFP.Columns(1).Value
            
                    If cCodFHijo = "001" Then 'SOLES'
                        frm_VTA_FormaPagoEfectivo.lblTipoCambio.Caption = cTC
                        frm_VTA_FormaPagoEfectivo.lblTipoCambio.BackColor = RGB(242, 242, 242)
                        frm_VTA_FormaPagoEfectivo.lblTipoCambio.ForeColor = RGB(0, 0, 0)
                        frm_VTA_FormaPagoEfectivo.LblTituloImporte.Caption = "Importe S/. : "
                        frm_VTA_FormaPagoEfectivo.txtPagaCon.Text = cImporte
                    ElseIf cCodFHijo = "002" Then 'DOLARES'
                        'frm_VTA_FormaPagoEfectivo.pstrDato = objVenta.FormaPago(GrdListaFP.Row, 0)
                        'frm_VTA_FormaPagoEfectivo.cCodFPadre = objVenta.FormaPago(GrdListaFP.Row, 0)
                        frm_VTA_FormaPagoEfectivo.lblTipoCambio.BackColor = RGB(227, 255, 213)
                        frm_VTA_FormaPagoEfectivo.lblTipoCambio.ForeColor = RGB(0, 0, 0)
                        frm_VTA_FormaPagoEfectivo.LblTituloImporte.Caption = "Importe $. : "
                        If frm_VTA_ConsultaDoc.pblnFpago = False Then
                            frm_VTA_FormaPagoEfectivo.txtPagaCon.Text = (cImporte / cTC)
                            frm_VTA_FormaPagoEfectivo.lblImporte.Caption = cImporte
                          Else
                            frm_VTA_FormaPagoEfectivo.txtPagaCon.Text = cImporte
                            frm_VTA_FormaPagoEfectivo.lblImporte.Caption = cImporte * cTC
                        End If
                    End If
                    
                    'frm_VTA_FormaPagoEfectivo.Show
            frm_VTA_FormaPagoEfectivo.grdEfectivo.DataSource.FindFirst "COD_HIJO='" & Trim(cCodFHijo) & "'"
            ElseIf cCodFPadre = "002" Then 'TARJETAS'
                        frm_VTA_FormaPagoTarjeta.pstrDato = objVenta.FormaPago(GrdListaFP.row, 0)
                         frm_VTA_FormaPagoTarjeta.pstrDatoDes = objVenta.FormaPago(GrdListaFP.row, 1)
                        'frm_VTA_FormaPagoTarjeta.cCodFPadre = objVenta.FormaPago(GrdListaFP.Row, 0)
                        frm_VTA_FormaPagoTarjeta.Show
                        '**************************************************************************************'
                        '** Creado 04/10/2007 Por Cristhian Rueda **'
                        If objUsuario.EsDelivery Then
                            frm_VTA_FormaPagoTarjeta.txtNomTitular.Text = objVenta.NomTitular
                            frm_VTA_FormaPagoTarjeta.txtNumDNI.Text = objVenta.NumDNI
                            frm_VTA_FormaPagoTarjeta.mskVencimiento.Text = Format(objVenta.FormaPago(GrdListaFP.row, 14), "mm/yyyy")
                        End If
                        '**************************************************************************************'
                        'I.ECASTILLO 07.10.2020
                        'If objVenta.FormaPago(GrdListaFP.row, 12) = "4100000000000000" Then frm_VTA_FormaPagoTarjeta.ctlCboTipoTarjeta.BoundText = "003"
                        'If objVenta.FormaPago(GrdListaFP.row, 12) = "5100000000000000" Then frm_VTA_FormaPagoTarjeta.ctlCboTipoTarjeta.BoundText = "004"
                        frm_VTA_FormaPagoTarjeta.ctlCboTipoTarjeta.BoundText = "" & gclsOracle.FN_Valor("BTLPROD.PKG_FORMA_PAGO.FN_GET_HIJO", objVenta.FormaPago(GrdListaFP.row, 12))
                        'F.ECASTILLO 07.10.2020
                        frm_VTA_FormaPagoTarjeta.txtNroTar.Text = objVenta.FormaPago(GrdListaFP.row, 12)
                        frm_VTA_FormaPagoTarjeta.LblNomTraj.Caption = objTarjeta.ValidaTarjeta(frm_VTA_FormaPagoTarjeta.txtNroTar.Text, "2")
                        
                        '26/08/07 comentado por Pherrera
                        'frm_VTA_FormaPagoTarjeta.mskVencimiento.Text = Format(objVenta.FormaPago(GrdListaFP.Row, 14), "mm/yyyy")
                        frm_VTA_FormaPagoTarjeta.txtNroCuota.Text = objVenta.FormaPago(GrdListaFP.row, 13)
                        frm_VTA_FormaPagoTarjeta.txtNroAut.Text = objVenta.FormaPago(GrdListaFP.row, 20)
                        frm_VTA_FormaPagoTarjeta.ctLCboTipoCuota.BoundText = objVenta.FormaPago(GrdListaFP.row, 15)
                        frm_VTA_FormaPagoTarjeta.txtImporte.Text = objVenta.FormaPago(GrdListaFP.row, 4)
                        
                        frm_VTA_FormaPagoTarjeta.ChkRetEfec.Value = 0
                        frm_VTA_FormaPagoTarjeta.TxtRetiro.Text = ""
                        
                        Dim rs As oraDynaset
                        Set rs = objFormaPago.ListaFormaPDocumento(Param_Tipo_Documento, Param_Numero_Documento)
                        While Not rs.EOF
                            If rs("COD_FORMA_PAGO").Value = cCodFPadre And rs("COD_HIJO").Value = objVenta.FormaPago(GrdListaFP.row, 2) And rs("NUM_TARJETA").Value = objVenta.FormaPago(GrdListaFP.row, 12) Then
                                If rs("FLG_RETIRO_EFEC").Value = "1" Then
                                    frm_VTA_FormaPagoTarjeta.ChkRetEfec.Value = 1
                                    frm_VTA_FormaPagoTarjeta.TxtRetiro.Text = rs("IMP_SIN_REDONDEO").Value
                                End If
                            End If
                            rs.MoveNext
                        Wend
                        
'                        If objVenta.FormaPago(GrdListaFP.Row, 29) <> "" And objVenta.FormaPago(GrdListaFP.Row, 29) > 0 Then
'                            frm_VTA_FormaPagoTarjeta.ChkRetEfec.Value = 1
'                            frm_VTA_FormaPagoTarjeta.TxtRetiro.Text = objVenta.FormaPago(GrdListaFP.Row, 29)
'                          Else
'                            frm_VTA_FormaPagoTarjeta.ChkRetEfec.Value = 0
'                            frm_VTA_FormaPagoTarjeta.TxtRetiro.Text = ""
'                        End If
                        
            ElseIf cCodFPadre = "003" Then 'CREDITO'
                        frm_VTA_FormaPagoCredito.pstrDato = objVenta.FormaPago(GrdListaFP.row, 0)
                        'frm_VTA_FormaPagoCredito.cCodFPadre = objVenta.FormaPago(GrdListaFP.Row, 0)
                        frm_VTA_FormaPagoCredito.Show
                        frm_VTA_FormaPagoCredito.txtValor.Text = objVenta.FormaPago(GrdListaFP.row, 4)
                        
            ElseIf cCodFPadre = "004" Then 'NOTA CREDITO'
                        frm_VTA_FormaPagoNC.pstrDato = objVenta.FormaPago(GrdListaFP.row, 0)
                        'frm_VTA_FormaPagoNC.cCodFPadre = objVenta.FormaPago(GrdListaFP.Row, 0)
                        frm_VTA_FormaPagoNC.Show
                        frm_VTA_FormaPagoNC.txtNroNC.Text = objVenta.FormaPago(GrdListaFP.row, 22)
                        frm_VTA_FormaPagoNC.lblFecha.Caption = objVenta.FormaPago(GrdListaFP.row, 21)
                        frm_VTA_FormaPagoNC.lblLocal.Caption = objVenta.FormaPago(GrdListaFP.row, 27)
                        frm_VTA_FormaPagoNC.lblImporte.Caption = objVenta.FormaPago(GrdListaFP.row, 4)
                        frm_VTA_FormaPagoNC.lblReferencia.Caption = objVenta.FormaPago(GrdListaFP.row, 28)
                                    
            ElseIf cCodFPadre = "006" Then 'CHEQUE'
                  frm_VTA_FormaPagoCheque.pstrDato = objVenta.FormaPago(GrdListaFP.row, 0)
                  'frm_VTA_FormaPagoCheque.cCodFPadre = objVenta.FormaPago(GrdListaFP.Row, 0)
                  frm_VTA_FormaPagoCheque.txtNroChq.Text = objVenta.FormaPago(GrdListaFP.row, 23)
                  frm_VTA_FormaPagoCheque.mskFecEmi.Text = objVenta.FormaPago(GrdListaFP.row, 21)
                  frm_VTA_FormaPagoCheque.ctlCboMoneda.BoundText = objVenta.FormaPago(GrdListaFP.row, 6)
                  frm_VTA_FormaPagoCheque.lblTipoCambio.Caption = objVenta.FormaPago(GrdListaFP.row, 11)
                  If objVenta.FormaPago(GrdListaFP.row, 6) = "1" Then
                        frm_VTA_FormaPagoCheque.txtMonto.Text = objVenta.FormaPago(GrdListaFP.row, 4)
                        frm_VTA_FormaPagoCheque.lblImporte.Caption = objVenta.FormaPago(GrdListaFP.row, 4)
                     Else
                        frm_VTA_FormaPagoCheque.txtMonto.Text = (cImporte / cTC) 'objVenta.FormaPago(GrdListaFP.Row, 4)
                        frm_VTA_FormaPagoCheque.lblImporte.Caption = objVenta.FormaPago(GrdListaFP.row, 4)
                  End If
                  
             ElseIf cCodFPadre = "007" Then 'DOCUMENTO DESCUENTO'
                 'Unload frm_VTA_FormaPagoDD
                  frm_VTA_FormaPagoDD.pstrDato = objVenta.FormaPago(GrdListaFP.row, 0)
                  'frm_VTA_FormaPagoDD.cCodFPadre = objVenta.FormaPago(GrdListaFP.Row, 0)
                  frm_VTA_FormaPagoDD.txtNroDoc.Text = objVenta.FormaPago(GrdListaFP.row, 16)
                  frm_VTA_FormaPagoDD.mskFecEmi.Text = objVenta.FormaPago(GrdListaFP.row, 21)
                  frm_VTA_FormaPagoDD.txtNombre.Text = objVenta.FormaPago(GrdListaFP.row, 24)
                  frm_VTA_FormaPagoDD.txtDNI.Text = objVenta.FormaPago(GrdListaFP.row, 25)
                  frm_VTA_FormaPagoDD.txtValor.Text = objVenta.FormaPago(GrdListaFP.row, 4)
                  frm_VTA_FormaPagoDD.od.FindFirst "COD_HIJO=" & "'" & objVenta.FormaPago(GrdListaFP.row, 2) & "'"
                  'frm_VTA_FormaPagoDD.Show
                  frm_VTA_FormaPagoDD.SetFocus
            ElseIf cCodFPadre = "011" Then
                frm_VTA_FormaPagoVF.pstrDato = objVenta.FormaPago(GrdListaFP.row, 0)
                frm_VTA_FormaPagoVF.txtNumero.Text = objVenta.FormaPago(GrdListaFP.row, 34)
                frm_VTA_FormaPagoVF.mskFec.Text = objVenta.FormaPago(GrdListaFP.row, 21)
                frm_VTA_FormaPagoVF.txtValor.Text = objVenta.FormaPago(GrdListaFP.row, 4)
                frm_VTA_FormaPagoVF.Show
            End If
                
        Exit Sub
CtlrErr:
        MsgBox Err.Description, vbCritical, App.FileDescription
            
End Sub
