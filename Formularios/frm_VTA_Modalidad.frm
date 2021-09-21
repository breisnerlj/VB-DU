VERSION 5.00
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_VTA_Modalidad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modalidad de venta"
   ClientHeight    =   6780
   ClientLeft      =   6195
   ClientTop       =   615
   ClientWidth     =   4725
   ControlBox      =   0   'False
   Icon            =   "frm_VTA_Modalidad.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   4725
   StartUpPosition =   2  'CenterScreen
   Begin TrueDBGrid70.TDBGrid grdModalidad 
      Bindings        =   "frm_VTA_Modalidad.frx":000C
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   11880
      _LayoutType     =   4
      _RowHeight      =   31
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Codigo"
      Columns(0).DataField=   "COD_MODALIDAD_VENTA"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descripción"
      Columns(1).DataField=   "DES_MENU"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   528
      Columns(2)._MaxComboItems=   5
      Columns(2).ValueItems(0)._DefaultItem=   0
      Columns(2).ValueItems(0).Value=   "001"
      Columns(2).ValueItems(0).Value.vt=   8
      Columns(2).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(2).ValueItems(0).DisplayValue(0)=   "bHQAADYOAABCTTYOAAAAAAAANgAAACgAAAAlAAAAIAAAAAEAGAAAAAAAAA4AAAAAAAAAAAAAAAAA"
      Columns(2).ValueItems(0).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(2)=   "///////////////////37+fWz8b37+f/////////////////////////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(3)=   "//8A////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(4)=   "////rf//td/WWq6lIbatKZaMOWFKnIZz////////////////////////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(5)=   "AP///////////////////////////////////////////////////////////////////7Xf1jnn"
      Columns(2).ValueItems(0).DisplayValue(6)=   "7wD3/wD3/wjn7wD3/wD3/zmOc0ogALWWc////////////////////////////////////////wD/"
      Columns(2).ValueItems(0).DisplayValue(7)=   "///////////////////////////////////////////////////////37+/OtqU5YUo5YUoI5+8A"
      Columns(2).ValueItems(0).DisplayValue(8)=   "9/8I5+8A9/8I5+8A9/85aVprKABrQSH37+f///////////////////////////////////8A////"
      Columns(2).ValueItems(0).DisplayValue(9)=   "////////////////////////////////////////////1se1nIZzaygASjAQa0EYOWlaOefvAPf/"
      Columns(2).ValueItems(0).DisplayValue(10)=   "APf/APf/APf/APf/OY5zaygAaygA1s/G////////////////////////////////////AP//////"
      Columns(2).ValueItems(0).DisplayValue(11)=   "//////////////////////////////fv55yGc5RhOVogAGsoAFogAGsoAGtBITlpWgjn7wD3/wjn"
      Columns(2).ValueItems(0).DisplayValue(12)=   "7wjf5wjn7wD3/zlpWmsoAFogAFpZMf///////////////////////////////////wD/////////"
      Columns(2).ValueItems(0).DisplayValue(13)=   "///////////////////Wz8aUYTlKIABrKABrKAB7OAhrKABrKABrKABaWTE5aVoA9/8A9/8A9/8I"
      Columns(2).ValueItems(0).DisplayValue(14)=   "5+8A9/8A9/+chnNaIABrKABaIAD37+f///////////////////////////////8A////////////"
      Columns(2).ValueItems(0).DisplayValue(15)=   "////////1s/Ga0EhWiAAWiAAaygAWiAAaygAYzAQaygAa0EhaygAa0EhOWFKCOfvCN/nCOfvCN/n"
      Columns(2).ValueItems(0).DisplayValue(16)=   "COfvAPf/MYZ7aygAa0EhaygAnIZz////////////////////////////////AP//////////////"
      Columns(2).ValueItems(0).DisplayValue(17)=   "/+ff1msoAGsoAGsoAGsoAGsoAEKGaxDPzimupRi+tSmupRi+tSmupRDPzgD3/wD3/wD3/wjn7wD3"
      Columns(2).ValueItems(0).DisplayValue(18)=   "/wD3/xjHxiGmnCmupSGmnFqupf///////////////////////////////wD///////////////+c"
      Columns(2).ValueItems(0).DisplayValue(19)=   "hnNrKABrQSFrKABaIAAprqUI5+8A9/8A9/8A9/8I5+8A9/8A9/8A9/8I5+8A9/8I5+8A9/8I5+8A"
      Columns(2).ValueItems(0).DisplayValue(20)=   "9/8A9/8A9/8I5+8A9/8Qz87Wz8b///////////////////////////8A////////////////lGE5"
      Columns(2).ValueItems(0).DisplayValue(21)=   "SjAQaygAaygAezgIOWlaOY5zAPf/OefvCOfvAPf/COfvAPf/COfvCN/nAPf/OefvCOfvAPf/APf/"
      Columns(2).ValueItems(0).DisplayValue(22)=   "OefvCOfvAPf/APf/CN/nWr69////////////////////////////AP///////////////5yGc2tB"
      Columns(2).ValueItems(0).DisplayValue(23)=   "GFogAGsoAFogADmOczGGewD3/wjn7wD3/wjn7wjf5wjn7wD3/wjn7wjf5wjn7wD3/wjn7wjf5wjn"
      Columns(2).ValueItems(0).DisplayValue(24)=   "7wD3/wjn7wjf5wjn7zmOc/fv7////////////////////////wD////////////////Wx7WUcVqE"
      Columns(2).ValueItems(0).DisplayValue(25)=   "SSFrKABrKAA5aVopppQA9/8A9/8A9/8A9/8A9/8A9/8A9/8A9/8I5+8A9/8A9/8A9/8A9/8A9/8A"
      Columns(2).ValueItems(0).DisplayValue(26)=   "9/8A9/8A9/8A9/8Qz87OtqX///////////////////////8A////////////////////lGE5UmFK"
      Columns(2).ValueItems(0).DisplayValue(27)=   "aygAaygAOY5zOWlaGMfGEM/OKa6lGL61Ka6lGL61GMfGCOfvAPf/COfvCN/nCOfvCN/nCOfvEM/O"
      Columns(2).ValueItems(0).DisplayValue(28)=   "COfvGMfGEM/OGMfGOWla////////////////////////AP///////////////////62WhJRhOWso"
      Columns(2).ValueItems(0).DisplayValue(29)=   "AHs4CBi+tWtBGGsoAFpZMWtBIWtBGEJBIXs4CGsoAAD3/wD3/wD3/wjn7wD3/wjn73tpWkJBIVpZ"
      Columns(2).ValueItems(0).DisplayValue(30)=   "MUJBIXs4CEJBIWsoANbPxv///////////////////wD////////////////////Wz8aEUSGESSFa"
      Columns(2).ValueItems(0).DisplayValue(31)=   "IAA5YUpCQSFaWTE5YUpCQSFrQSFaWTExhns5YUoI5+8A9/8I5+8A9/8I5+8I3+c5aVprKABrQSFa"
      Columns(2).ValueItems(0).DisplayValue(32)=   "WTFCQSFaWTFaIACchnP///////////////////8A////////////////////////lHFae2laaygA"
      Columns(2).ValueItems(0).DisplayValue(33)=   "SjAQezgIaygAaygAaygAaygASiAAa0EYOWFKOefvAPf/APf/COfvAPf/COfvWlkxaygAaygASjAQ"
      Columns(2).ValueItems(0).DisplayValue(34)=   "aygAaygAaygAa0Eh9+/n////////////////AP///////////////////////86+tYRJIWtBIWso"
      Columns(2).ValueItems(0).DisplayValue(35)=   "AGsoAGsoAGsoAGsoAGsoAGsoAGtBIVpZMQjn7wD3/wjn7wjf5wjn7wjf52tBIWsoAGsoAGsoAGso"
      Columns(2).ValueItems(0).DisplayValue(36)=   "AGsoAGtBIWsoAM6+tf///////////////wD///////////////////////////9aWTGUYTlrKABr"
      Columns(2).ValueItems(0).DisplayValue(37)=   "KABrKAB7OAhKMBBrKABrKABaWTE5aVoA9/8A9/8A9/8I5+8A9/8I5+9aWTFrKABrKABrKAB7OAhr"
      Columns(2).ValueItems(0).DisplayValue(38)=   "KAB7OAhrKACUYTne//////////////8A////////////////////////////tZZzhGFKaygAaygA"
      Columns(2).ValueItems(0).DisplayValue(39)=   "aygAYzAQaygAaygAaygAa0EhWlkxCOfvCN/nCOfvEM/OCOfvCN/na0EhaygAaygAaygAYzAQaygA"
      Columns(2).ValueItems(0).DisplayValue(40)=   "a0EhaygAYzAQ9+/n////////////AP///////////////////////////9bPxpyGc2soAGsoAGso"
      Columns(2).ValueItems(0).DisplayValue(41)=   "AIRJIWsoAGsoAGsoAFpZMTlhSgD3/wD3/wD3/wD3/wD3/wjn71pZMWsoAHs4CGsoAGsoAGsoAHs4"
      Columns(2).ValueItems(0).DisplayValue(42)=   "CEowEGsoANbPxv///////////wD////////////////////////////37+eUcVqUYTlaIABrKABr"
      Columns(2).ValueItems(0).DisplayValue(43)=   "KABrKABrKABrKABrQSE5YUo55+8I3+cI5+8I3+cI5+8I3+drQSFrKABrKABrKABrKABrKABrQSFr"
      Columns(2).ValueItems(0).DisplayValue(44)=   "KABrQSH37+f///////////8A////////////////////////////////tZZze2laaygAaygAezgI"
      Columns(2).ValueItems(0).DisplayValue(45)=   "aygAaygAaygAWlkxOWlaWlkxaygAaygAa0EhWlkxOWFKezgIaygAaygASjAQaygAaygAa0EYaygA"
      Columns(2).ValueItems(0).DisplayValue(46)=   "zral////////////////AP////////////////////////////////fv72soAIRRIWsoAGtBIWso"
      Columns(2).ValueItems(0).DisplayValue(47)=   "AGsoAHs4CGtBITmOczlpWjmOczGGezmOczlpWmsoAFogAGsoAGtBIWsoAFogAGsoAJRxWtbPxv//"
      Columns(2).ValueItems(0).DisplayValue(48)=   "/////////////////wD///////////////////////////////////9aWTGUYTlrKACESSFrKABr"
      Columns(2).ValueItems(0).DisplayValue(49)=   "KABrKABrKABrKAB7OAhrKABrKABrKABrKABKMBCESSFKMBBrKABaWTHOtqXWz8b/////////////"
      Columns(2).ValueItems(0).DisplayValue(50)=   "//////////////8A////////////////////////////////////9+/nhGFKaygAWiAAaygAWiAA"
      Columns(2).ValueItems(0).DisplayValue(51)=   "aygAaygAaygAWiAAaygAaygAaygAWiAAaygAa0EhlGE57+fe////////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(52)=   "////////////AP///////////////////////////////////97//5yGc1pZMXs4CGsoAGsoAGso"
      Columns(2).ValueItems(0).DisplayValue(53)=   "AHs4CGsoAIRJIWsoAGsoAGtBIZRhOa2WhP//////////////////////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(54)=   "/////////wD///////////////////////////////////////+chnOccVJrQSFrKABaIABrKABr"
      Columns(2).ValueItems(0).DisplayValue(55)=   "QSFrKABrKABrKACchnPWz8b37+//////////////////////////////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(56)=   "//////8A////////////////////////////////////////9+/nrZaEtZZza0EhhEkha0EhaygA"
      Columns(2).ValueItems(0).DisplayValue(57)=   "a0Ehzral3v//////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(58)=   "////AP////////////////////////////////////////////fv562WhJRhOYRRIZyGc/fv7///"
      Columns(2).ValueItems(0).DisplayValue(59)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(60)=   "/wD////////////////////////////////////////////////////e////////////////////"
      Columns(2).ValueItems(0).DisplayValue(61)=   "//////////////////////////////////////////////////////////////////////////8A"
      Columns(2).ValueItems(0).DisplayValue(62)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(63)=   "////////////////////////////////////////////////////////////////////////AA=="
      Columns(2).ValueItems(0).DisplayValue.vt=   9
      Columns(2).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(1)._DefaultItem=   0
      Columns(2).ValueItems(1).Value=   "002"
      Columns(2).ValueItems(1).Value.vt=   8
      Columns(2).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(2).ValueItems(1).DisplayValue(0)=   "bHQAADYOAABCTTYOAAAAAAAANgAAACgAAAAlAAAAIAAAAAEAGAAAAAAAAA4AAAAAAAAAAAAAAAAA"
      Columns(2).ValueItems(1).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(1).DisplayValue(2)=   "///////////////////////Wz8b37+f/////////////////////////////////////////////"
      Columns(2).ValueItems(1).DisplayValue(3)=   "//8A////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(1).DisplayValue(4)=   "////9+/nxqaUlGE5a0EpYygAa0EpnIZz9+/v////////////////////////////////////////"
      Columns(2).ValueItems(1).DisplayValue(5)=   "AP////////////////////////////////////////////////////////////////////f375yG"
      Columns(2).ValueItems(1).DisplayValue(6)=   "c2tBKWMoAGMoAGMoAGMoAGMoAEogAM6+rf///////////////////////////////////////wD/"
      Columns(2).ValueItems(1).DisplayValue(7)=   "///////////////////////////////////////////////////////37+/Ovq1rQSlKIABaIAA5"
      Columns(2).ValueItems(1).DisplayValue(8)=   "jnsA9/8p185aIABjKABjKAAQz9Yhppz37+f///////////////////////////////////8A////"
      Columns(2).ValueItems(1).DisplayValue(9)=   "////////////////////////////////////////////1se1lLala0EYSiAAYygAYygAY4ZrAPf/"
      Columns(2).ValueItems(1).DisplayValue(10)=   "APf/IaacezgIYygAY4ZrAPf/EM/W1s/G////////////////////////////////////AP//////"
      Columns(2).ValueItems(1).DisplayValue(11)=   "//////////////////////////////fv55S2pZRhOVogAGMoAFogAGMoAFogAGMoAADn7wD3/wDn"
      Columns(2).ValueItems(1).DisplayValue(12)=   "7xDP1mtBKTmOewDn7wD3/wDn72OGa////////////////////////////////////wD/////////"
      Columns(2).ValueItems(1).DisplayValue(13)=   "///////////////////Wz8aUYTlKIABjKABjKAB7OAhKIABjKABjKAB7OAgYvrUA9/8A9/8A9/8A"
      Columns(2).ValueItems(1).DisplayValue(14)=   "9/8A9/8A9/8A9/8A5+8A9/8hppzn39b///////////////////////////////8A////////////"
      Columns(2).ValueItems(1).DisplayValue(15)=   "////////1s/Ga0EpSiAAWiAAYygAYygAYygAWjAQYygAYygAYygAMXFjAPf/AOfvCN/nAOfvCN/n"
      Columns(2).ValueItems(1).DisplayValue(16)=   "AOfvCN/nAOfvAPf/AOfvCN/nlHFa////////////////////////////////AP//////////////"
      Columns(2).ValueItems(1).DisplayValue(17)=   "/+ff1mMoAGMoAGMoAGMoAGMoAGMoAGMoAGMoAGMoAGMoAGMoAFpZMQD3/wD3/wD3/wD3/wD3/wD3"
      Columns(2).ValueItems(1).DisplayValue(18)=   "/wDn7wD3/wD3/wD3/wD3/2OGa/fv7////////////////////////////wD///////////////+l"
      Columns(2).ValueItems(1).DisplayValue(19)=   "jntjKABrQSljKABjKABjKABaIABjKABrQSljKABjKABjKABjKABaWTEA5+8A9/8A5+8A9/8A5+8A"
      Columns(2).ValueItems(1).DisplayValue(20)=   "9/8A5+8A9/8A5+8A9/8ptq3Wx7X///////////////////////////8A////////////////lGE5"
      Columns(2).ValueItems(1).DisplayValue(21)=   "SiAAYygAYygAYygAa0EpYygAYygAezgISiAAYygAa0EpYygAYygAezgIMaaUCO/3AOfvAPf/APf/"
      Columns(2).ValueItems(1).DisplayValue(22)=   "CO/3AOfvAPf/APf/KdfOWp6M////////////////////////////AP///////////////6WOe1pZ"
      Columns(2).ValueItems(1).DisplayValue(23)=   "MVogAGMoAFogAGMoAGMoAGMoAFowEGMoAGMoAGMoAFogAGMoAFogAGMoAADn7wD3/wDn7wD3/wDn"
      Columns(2).ValueItems(1).DisplayValue(24)=   "7wD3/wDn7wjf5wDn7zFxY+/n3v///////////////////////wD////////////////n39ZaaUqE"
      Columns(2).ValueItems(1).DisplayValue(25)=   "USFjKABjKABjKAB7OAhjKACEUSFjKABjKABjKABjKABjKABaWTExppQI7/cA9/8p3+cA9/8A9/8A"
      Columns(2).ValueItems(1).DisplayValue(26)=   "9/8A9/8A9/8A9/8xcWPWx7X///////////////////////8A////////////////////lGE5WmlK"
      Columns(2).ValueItems(1).DisplayValue(27)=   "QlE5YygAYygAWiAAYygAWiAAOY57MXFjYygAa0EpMaaUAPf/APf/APf/OY57WiAAWlkxEM/WAPf/"
      Columns(2).ValueItems(1).DisplayValue(28)=   "AOfvAPf/EM/Wa0EYa0Ep////////////////////////AP///////////////////5S2pSnf5wD3"
      Columns(2).ValueItems(1).DisplayValue(29)=   "/ym2rWMoAGMoAFogAGOGawD3/ynXzmMoAIRRIQDn7wD3/wD3/wD3/zGmlGMoAGMoAGMoABDP1gD3"
      Columns(2).ValueItems(1).DisplayValue(30)=   "/wD3/2tBGFogAGMoANbPxv///////////////////wD////////////////////Wx7Uptq0A9/8A"
      Columns(2).ValueItems(1).DisplayValue(31)=   "9/9jKABaIAA5jnsA5+8A9/8A9/8xppRaIAAxppQA9/8A9/8A5+8p185aIABjKABaMBBjKAAA5+85"
      Columns(2).ValueItems(1).DisplayValue(32)=   "jntaIABjKABaIACchnP///////////////////8A////////////////////////lGE5APf/Kd/n"
      Columns(2).ValueItems(1).DisplayValue(33)=   "WlkxKbatAPf/APf/APf/APf/APf/KbatWiAAMaaUAPf/APf/MXFjYygAYygAYygAYygAYygASiAA"
      Columns(2).ValueItems(1).DisplayValue(34)=   "YygAYygAYygAa0Ep9+/n////////////////AP///////////////////////9bPxjmOewD3/wjf"
      Columns(2).ValueItems(1).DisplayValue(35)=   "5wDn7wD3/wDn7wD3/wDn7wD3/wDn72MoAFogAFpZMUJZQmMoAFogAGMoAGMoAGMoAGMoAGMoAGMo"
      Columns(2).ValueItems(1).DisplayValue(36)=   "AGMoAGtBKWMoAM6+tf///////////////wD////////////////////////37+cxppQA9/8A9/8A"
      Columns(2).ValueItems(1).DisplayValue(37)=   "9/8A9/8A9/8Qz9ZaWTFCUTlaWTFjKAB7OAhjKAB7OAhjKABjKABjKAB7OAhKIABjKABjKAB7OAhj"
      Columns(2).ValueItems(1).DisplayValue(38)=   "KABjKABKIACchnP39+////////////8A////////////////////////////nIZzGL61CN/nAOfv"
      Columns(2).ValueItems(1).DisplayValue(39)=   "CN/nAOfvOY57WiAAYygAWiAAYygAGL61WlkxWiAAYygAYygAYygAWiAAYygAYygAYygAWjAQYygA"
      Columns(2).ValueItems(1).DisplayValue(40)=   "a0EpYygAWjAQ9+/n////////////AP///////////////////////////9bPxim2rQD3/wD3/wD3"
      Columns(2).ValueItems(1).DisplayValue(41)=   "/wD3/zGmlHs4CGMoAGMoADmOewD3/wD3/1pZMWMoAGMoAGMoAGMoAGMoAGMoAGMoAGMoAGMoAHs4"
      Columns(2).ValueItems(1).DisplayValue(42)=   "CEogAGMoANbPxv///////////wD////////////////////////////37+danowA9/8A5+8A9/8A"
      Columns(2).ValueItems(1).DisplayValue(43)=   "9/8ptq1aIABjKABaIABaWTEA5+8A9/8A5+9jKABaIABjKABjKABjKABjKABjKABjKABjKABjKABj"
      Columns(2).ValueItems(1).DisplayValue(44)=   "KABrQSn37+f///////////8A////////////////////////////////tZZ7IaacOY57AOfvAPf/"
      Columns(2).ValueItems(1).DisplayValue(45)=   "APf/Y4ZrSiAAWlkxEM/WAPf/APf/APf/GL61YygAYygAYygASiAAYygASiAAYygAYygAa0EYYygA"
      Columns(2).ValueItems(1).DisplayValue(46)=   "zr6t////////////////AP////////////////////////////////fv73s4CIxRIWtBGADn7wD3"
      Columns(2).ValueItems(1).DisplayValue(47)=   "/wD3/wD3/wD3/wD3/wDn7wD3/wDn7wD3/0JBKWMoAFowEGMoAGtBKXs4CFogAGMoAJRxWtbPxv//"
      Columns(2).ValueItems(1).DisplayValue(48)=   "/////////////////wD///////////////////////////////////9aWTGUYTlaIAAp184A9/8A"
      Columns(2).ValueItems(1).DisplayValue(49)=   "9/8A5+8A9/8A9/8A9/8A9/8A9/8xcWNjKABKIACUYTlKIABjKABaWTHOvq3Wz8b/////////////"
      Columns(2).ValueItems(1).DisplayValue(50)=   "//////////////8A////////////////////////////////////9+/njFEhYygAKbatAPf/AOfv"
      Columns(2).ValueItems(1).DisplayValue(51)=   "CN/nAOfvCN/nAOfvCN/nWmlKYygAWiAAYygAa0EplGE57+fe////////////////////////////"
      Columns(2).ValueItems(1).DisplayValue(52)=   "////////////AP////////////////////////////////////f375RhOTFxYwD3/wDn7wD3/wD3"
      Columns(2).ValueItems(1).DisplayValue(53)=   "/wD3/wD3/ynXzmMoAGMoAGtBKZRhOZS2pf//////////////////////////////////////////"
      Columns(2).ValueItems(1).DisplayValue(54)=   "/////////wD///////////////////////////////////////+UtqUp184A5+8A9/8A5+8A9/8A"
      Columns(2).ValueItems(1).DisplayValue(55)=   "5+85jntjKABjKACljnvOvq3/////////////////////////////////////////////////////"
      Columns(2).ValueItems(1).DisplayValue(56)=   "//////8A////////////////////////////////////////9+/nlHFaKd/nAPf/APf/AOfvOY57"
      Columns(2).ValueItems(1).DisplayValue(57)=   "a0Epzr6t9/fv////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(1).DisplayValue(58)=   "////AP////////////////////////////////////////////fv56WOe5RhOWtBKZRpQuff1v//"
      Columns(2).ValueItems(1).DisplayValue(59)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(1).DisplayValue(60)=   "/wD////////////////////////////////////////////////////39+//////////////////"
      Columns(2).ValueItems(1).DisplayValue(61)=   "//////////////////////////////////////////////////////////////////////////8A"
      Columns(2).ValueItems(1).DisplayValue(62)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(1).DisplayValue(63)=   "////////////////////////////////////////////////////////////////////////AA=="
      Columns(2).ValueItems(1).DisplayValue.vt=   9
      Columns(2).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(2)._DefaultItem=   0
      Columns(2).ValueItems(2).Value=   "003"
      Columns(2).ValueItems(2).Value.vt=   8
      Columns(2).ValueItems(2).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(2).ValueItems(2).DisplayValue(0)=   "bHQAADYOAABCTTYOAAAAAAAANgAAACgAAAAlAAAAIAAAAAEAGAAAAAAAAA4AAAAAAAAAAAAAAAAA"
      Columns(2).ValueItems(2).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(2).DisplayValue(2)=   "///////////////////v7+/e39bv7+f/////////////////////////////////////////////"
      Columns(2).ValueItems(2).DisplayValue(3)=   "//8A////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(2).DisplayValue(4)=   "////////xr6tjGlSYzAYWiAAWigQpY57////////////////////////////////////////////"
      Columns(2).ValueItems(2).DisplayValue(5)=   "AP////////////////////////////////////////////////////////////////////f375R5"
      Columns(2).ValueItems(2).DisplayValue(6)=   "a2MwKVogAGMgAGsoAHMoAGsoAFIYAKWOhP///////////////////////////////////////wD/"
      Columns(2).ValueItems(2).DisplayValue(7)=   "///////////////////////////////////////////////////////39/etx8Z7SSlSCABaGABr"
      Columns(2).ValueItems(2).DisplayValue(8)=   "KABzMABzMABzMABzMABzMABrKABjMBj39/f///////////////////////////////////8A////"
      Columns(2).ValueItems(2).DisplayValue(9)=   "////////////////////////////////////////////zse9c8/OCOfvAP//MXFSa0EQczAAczAA"
      Columns(2).ValueItems(2).DisplayValue(10)=   "czAAczAAczAAczAAczAAczAAShAAzr61////////////////////////////////////AP//////"
      Columns(2).ValueItems(2).DisplayValue(11)=   "/////////////////////////////+/v57WmlHtRQlogAEJJMTGWhBDXzgD//yG2rUKGa3sgAHMw"
      Columns(2).ValueItems(2).DisplayValue(12)=   "AHMwAHMwAHMwAHMwAHMwAFogAGtBKf///////////////////////////////////wD/////////"
      Columns(2).ValueItems(2).DisplayValue(13)=   "///////////////////e19aEaVJaIABaGABrIABzMAB7KAB7IABzKABjSSFCeWMA9/9ChnN7IABz"
      Columns(2).ValueItems(2).DisplayValue(14)=   "MABzMABzMABzMABzMABzMABKEAD39/f///////////////////////////////8A////////////"
      Columns(2).ValueItems(2).DisplayValue(15)=   "////////3tfGa0EhQgAAYxgAczAAczAAczAAczAAczAAczAAczAAcygAjBAASnFSGL69UmFCczAA"
      Columns(2).ValueItems(2).DisplayValue(16)=   "czAAczAAczAAczAAczAAWhgAnIZz////////////////////////////////AP//////////////"
      Columns(2).ValueItems(2).DisplayValue(17)=   "/97XzmMwEGMgAHMwAHMwAHMwAHMwAHMwAHMwAHMwAHMwAHMwAHMwAHsgAEKGa0pxUkpxWkppSmNJ"
      Columns(2).ValueItems(2).DisplayValue(18)=   "IXsoAHMwAHMwAHMwAGsoAHNJMff39////////////////////////////wD///////////////+l"
      Columns(2).ValueItems(2).DisplayValue(19)=   "hnNjGABzMABzMABzMABzMABzMABzMABzMABzMABzMABzMABzMAB7IAA5jnM5lns5lns5loRCeWNj"
      Columns(2).ValueItems(2).DisplayValue(20)=   "SSFzKABzMABzMABzMABaIADOvrX///////////////////////////8A////////////////nGlK"
      Columns(2).ValueItems(2).DisplayValue(21)=   "ezgIczAAczAAczAAczAAazgYazAIeygAeygAcygAczAAczAAeyAAOZZ7OYZrAPf/MZaEOY5zQoZr"
      Columns(2).ValueItems(2).DisplayValue(22)=   "azgYcygAczAAczAAWhgAlHlr////////////////////////////AP///////////////62Oa4xZ"
      Columns(2).ValueItems(2).DisplayValue(23)=   "MXMwAHMwAHMwAHMoABjHvVJpSkpxWlpZOWtBEHsoAHMoAIQYADGejEKGawjv7wjv70KGazmOczGW"
      Columns(2).ValueItems(2).DisplayValue(24)=   "hHMwAHMoAHMwAGsoAFooCPfv7////////////////////////wD////////////////Wvq2laUKE"
      Columns(2).ValueItems(2).DisplayValue(25)=   "MABzKAB7IABSaUoI395ChmsxnoxCeWMI1945hmtjSSFzKAApppRChmsA7+8A//8hvrVChms5jnMp"
      Columns(2).ValueItems(2).DisplayValue(26)=   "rqV7KABzKABzMABKCADGtq3///////////////////////8A////////////////////UratKb61"
      Columns(2).ValueItems(2).DisplayValue(27)=   "UllCa0Epa0EQY1EpSnFaMZ6EGMfGSnFSGMfGEM/OczAYKaacQoZrAO/3APf/AP//MZaEOY57MZ6E"
      Columns(2).ValueItems(2).DisplayValue(28)=   "Ka6leygAczAAYyAAa0Eh////////////////////////AP///////////////////86ehHuGazGW"
      Columns(2).ValueItems(2).DisplayValue(29)=   "hCmmlDGehEJ5a1pRMXMwAFpROWs4CCGupSmmnIwIACG2rUKGawDv9wD3/wD3/wD3/0pxWjGWhBjH"
      Columns(2).ValueItems(2).DisplayValue(30)=   "vTGejIQYAHMwAFIYAOff3v///////////////////wD////////////////////n39acYTGMOBBz"
      Columns(2).ValueItems(2).DisplayValue(31)=   "KABjSSlKcVIpppQI194xppxzKAB7KABrOAiMAAAYvr1ChmsA9/8A9/8A9/8A//8I5+daWTE5jnsA"
      Columns(2).ValueItems(2).DisplayValue(32)=   "9/9SaUqEGABSGACtnoz///////////////////8A////////////////////////lGFCnGlKczAA"
      Columns(2).ValueItems(2).DisplayValue(33)=   "cygAeygAhBgAhBAAWlk5EM/OIa6lSnFSY1EhEM/OSnFaAPf/APf/APf/APf/AP//GMfGQnljSnFS"
      Columns(2).ValueItems(2).DisplayValue(34)=   "AP//SnFacyAAYzAY9/f3////////////////AP///////////////////////97HvYRJIXtBEHMw"
      Columns(2).ValueItems(2).DisplayValue(35)=   "AHMwAHMwAHMwAHsgAHsYAEKGa0J5Y0pxWlpROUpxUgD//wD3/wD3/wD3/wD3/wD3/zmWhCmmnDGe"
      Columns(2).ValueItems(2).DisplayValue(36)=   "hAD//2tBGFoYAMa2pf///////////////wD////////////////////////39++caUKMUSFrKABz"
      Columns(2).ValueItems(2).DisplayValue(37)=   "MABzMABzMABzMAB7KABKaUoA7/cYz8Y5jnsQ19YA//8A9/8A9/8A9/8A9/8A//8Q19ZKcVoYvr1K"
      Columns(2).ValueItems(2).DisplayValue(38)=   "cVJzOAhaGACMaWP///////////////8A////////////////////////////vaaElGE5aygAczAA"
      Columns(2).ValueItems(2).DisplayValue(39)=   "czAAczAAczAAczAAeyAAOYZzAP//AP//AP//CN/nAOfvAPf/APf/AP//GL69OYZrMZaEY0kheyAA"
      Columns(2).ValueItems(2).DisplayValue(40)=   "czAAaygAYzAQ7+fn////////////AP///////////////////////////97PvaV5Wns4CHMoAHMw"
      Columns(2).ValueItems(2).DisplayValue(41)=   "AHMwAHMwAHMwAHMwAHsoADmWhAD//xDPziG+rUKGawD//wD39ymunDmOczmWe2s4AHsgAHMwAHMw"
      Columns(2).ValueItems(2).DisplayValue(42)=   "AHMoAFooAO/n5////////////wD////////////////////////////39++teWOUYTlrKABzMABz"
      Columns(2).ValueItems(2).DisplayValue(43)=   "MABzMABzMABzMABzKAB7GAAhtqUQz84hvrUxnowI3+cxloQpppQxppRaUSlzKABzMABzMABzMABj"
      Columns(2).ValueItems(2).DisplayValue(44)=   "IAB7STn39/f///////////8A////////////////////////////////tZZzpXlSczAAczAAczAA"
      Columns(2).ValueItems(2).DisplayValue(45)=   "czAAczAAczAAczAAcygAcygAENfWGM/GCO/vMZaEGL69OZZ7QoZjcygAczAAczAAczAAczAAYyAA"
      Columns(2).ValueItems(2).DisplayValue(46)=   "zr6t////////////////AP///////////////////////////////////4RJGIxZMWsoAHMwAHMw"
      Columns(2).ValueItems(2).DisplayValue(47)=   "AHMwAHMwAHMwAHMwAHsoAGNBGBDX1kJ5Yxi+vVJhOSmelIQgAHs4CHMwAHMwAIRJIZRpSuff1v//"
      Columns(2).ValueItems(2).DisplayValue(48)=   "/////////////////wD///////////////////////////////////+caUqUYTlzMABzMABzMABz"
      Columns(2).ValueItems(2).DisplayValue(49)=   "MABzMABzMABzMABzMABzMABrOBBrOBBaUSlzKAB7QRBzMABjGAB7SRjGppTe187/////////////"
      Columns(2).ValueItems(2).DisplayValue(50)=   "//////////////8A////////////////////////////////////59fOjFEpe0EQczAAczAAczAA"
      Columns(2).ValueItems(2).DisplayValue(51)=   "czAAczAAczAAczAAczAAczAAczAAezAAezgIczAAnHFK59/W////////////////////////////"
      Columns(2).ValueItems(2).DisplayValue(52)=   "////////////AP////////////////////////////////////f396V5WoRJIWsoAHMwAHMwAHMw"
      Columns(2).ValueItems(2).DisplayValue(53)=   "AHMwAHMwAHM4CHM4AHMwAHM4CJRhOb2mjPfv5///////////////////////////////////////"
      Columns(2).ValueItems(2).DisplayValue(54)=   "/////////wD///////////////////////////////////////+9poyleVpzMABrKABrKABzMAB7"
      Columns(2).ValueItems(2).DisplayValue(55)=   "OAhzOAhzMACESRitjnPWx73/////////////////////////////////////////////////////"
      Columns(2).ValueItems(2).DisplayValue(56)=   "//////8A////////////////////////////////////////7+/ntZZztY5zlFkxhEkYczAAaygA"
      Columns(2).ValueItems(2).DisplayValue(57)=   "jFkxxrac9/f3////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(2).DisplayValue(58)=   "////AP/////////////////////////////////////////////397WWe5xpSoRRKaV5Y/fv7///"
      Columns(2).ValueItems(2).DisplayValue(59)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(2).DisplayValue(60)=   "/wD/////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(2).DisplayValue(61)=   "//////////////////////////////////////////////////////////////////////////8A"
      Columns(2).ValueItems(2).DisplayValue(62)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(2).DisplayValue(63)=   "////////////////////////////////////////////////////////////////////////AA=="
      Columns(2).ValueItems(2).DisplayValue.vt=   9
      Columns(2).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(3)._DefaultItem=   0
      Columns(2).ValueItems(3).Value=   "004"
      Columns(2).ValueItems(3).Value.vt=   8
      Columns(2).ValueItems(3).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(2).ValueItems(3).DisplayValue(0)=   "bHQAADYOAABCTTYOAAAAAAAANgAAACgAAAAlAAAAIAAAAAEAGAAAAAAAAA4AAAAAAAAAAAAAAAAA"
      Columns(2).ValueItems(3).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(3).DisplayValue(2)=   "///////////////////v7+fe39bv7+f/////////////////////////////////////////////"
      Columns(2).ValueItems(3).DisplayValue(3)=   "//8A////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(3).DisplayValue(4)=   "//////fvzq6cjGlSYzAYYxgAaxAApYZz////////////////////////////////////////////"
      Columns(2).ValueItems(3).DisplayValue(5)=   "AP/////////////////////////////////////////////////////////////////////f1pxx"
      Columns(2).ValueItems(3).DisplayValue(6)=   "Y0phSjlpUmMgAGsgADGehBC+xkooCK2Ga////////////////////////////////////////wD/"
      Columns(2).ValueItems(3).DisplayValue(7)=   "////////////////////////////////////////////////////////7+fGppRjcWMQjoQA3+cA"
      Columns(2).ValueItems(3).DisplayValue(8)=   "//8hvq17KAB7KAAhx60A//8A9/dKcVr/7+f///////////////////////////////////8A////"
      Columns(2).ValueItems(3).DisplayValue(9)=   "//////////////////////////////////////////f31ratnI57SnFjEK6lAO//AP//AP//AP//"
      Columns(2).ValueItems(3).DisplayValue(10)=   "IbaleyAAeyAAIb6tAP//AP//AK6tzr61////////////////////////////////////AP//////"
      Columns(2).ValueItems(3).DisplayValue(11)=   "/////////////////////////////+/v572ejHNpUiGOewDHxgD//wD//wD3/wD3/wD3/wD//yG2"
      Columns(2).ValueItems(3).DisplayValue(12)=   "rXsgAHsgACG+rQD//wD3/wD3/1phUv///////////////////////////////////wD/////////"
      Columns(2).ValueItems(3).DisplayValue(13)=   "///////////////////e19aEaVJjIABCQTEA5+cA//8A//8A9/8A9/8A9/8A9/8A9/8A//8hvq17"
      Columns(2).ValueItems(3).DisplayValue(14)=   "IAB7IAAhvq0A//8A9/8A//8QhnP/39b///////////////////////////////8A////////////"
      Columns(2).ValueItems(3).DisplayValue(15)=   "////////3tfGazghQgAAWhgAeyAAWllCAPf/APf/APf/APf/APf/APf/APf/APf/AP//Ib6teyAA"
      Columns(2).ValueItems(3).DisplayValue(16)=   "eyAAIb6tAP//APf/APf/APf/rWlS////////////////////////////////AP//////////////"
      Columns(2).ValueItems(3).DisplayValue(17)=   "/97XzmMwEGMgAHMwAHMwAHsoAFpZQgD3/wD3/wD3/wD3/wD3/wD3/wD3/wD3/wD//yHHtXsgAHsg"
      Columns(2).ValueItems(3).DisplayValue(18)=   "ACG+rQD//wD3/wD3/wD//2OGc//n5////////////////////////////wD///////////////+l"
      Columns(2).ValueItems(3).DisplayValue(19)=   "hnNjGABzMABzMABzMAB7KABaWUIA9/cA9/8A9/8A9/8A9/8A9/8A9/8A9/8A//8Yz717IAB7IAAh"
      Columns(2).ValueItems(3).DisplayValue(20)=   "vq0A//8A9/8A9/8A//8Yrq3Wvq3///////////////////////////8A////////////////nHFS"
      Columns(2).ValueItems(3).DisplayValue(21)=   "ezgIczAAczAAczAAcygAWlE5AO/3APf/APf/APf/APf/APf/APf/APf/AP//GM/GeyAAeyAAKbal"
      Columns(2).ValueItems(3).DisplayValue(22)=   "AP//APf/APf/APf/ANfWjJaM////////////////////////////AP///////////////62Oa4xZ"
      Columns(2).ValueItems(3).DisplayValue(23)=   "MXMwAHMwAHMwAHMoAGNJKQjn5wD3/wD3/wD3/wD3/wD3/wD3/wD3/wD//ym2pYQQAHMoABjPvQD/"
      Columns(2).ValueItems(3).DisplayValue(24)=   "/wD3/wD3/wD3/wD3/zF5c/fn5////////////////////////wD////////////////Wvq2caUp7"
      Columns(2).ValueItems(3).DisplayValue(25)=   "OAhrKABzMABzMAB7KAAYz8YA//8A9/8A9/8A9/8A9/8A9/8A9/8A//9KeVpKaVII394A9/8A9/8A"
      Columns(2).ValueItems(3).DisplayValue(26)=   "9/8A9/8A9/8A//8ArqXGrqX///////////////////////8A////////////////////nHFSlGE5"
      Columns(2).ValueItems(3).DisplayValue(27)=   "aygAczAAczAAcygASnFjAP//APf/APf/APf/APf/APf/AP//Ice1KZ6UAP//AP//APf/APf/APf/"
      Columns(2).ValueItems(3).DisplayValue(28)=   "APf/APf/APf/AP//Y1E5////////////////////////AP///////////////////8amlJxpSnMw"
      Columns(2).ValueItems(3).DisplayValue(29)=   "AHMwAHMoAFJhSjmOczmOewD39wD//wD//wD//wD//xjHvUJ5awjn5wD3/wD3/wD3/wD3/wD3/wD3"
      Columns(2).ValueItems(3).DisplayValue(30)=   "/wD3/wD3/wD//xCmnO/Hvf///////////////////wD////////////////////n39aUYUKESSFr"
      Columns(2).ValueItems(3).DisplayValue(31)=   "KABzMABzMAAI5+cpppxrOBA5lnsI394Yx8ZKeVpKcVII3+cA//8A9/8A//8A//8A//8A//8A//8A"
      Columns(2).ValueItems(3).DisplayValue(32)=   "//8A//8A//8IrqW9hnP///////////////////8A////////////////////////lGFCnGlKczAA"
      Columns(2).ValueItems(3).DisplayValue(33)=   "czAAcygAhBgAGM/GCO/nSnFaa0EQa0EQMZaEAP//AP//AP//APf3ENfWMZaEcygAlAAAczAIOY57"
      Columns(2).ValueItems(3).DisplayValue(34)=   "SnFSeyAAhAAAWjAY9/f3////////////////AP///////////////////////97HvYRJIXtBEHMw"
      Columns(2).ValueItems(3).DisplayValue(35)=   "AHMwAHMoAHsYADmOcwjn7wD//wD3/zGmlFpZMTGejEKGa1pZOXMwAHsYAHsgAFpZORDf1gD3/wD3"
      Columns(2).ValueItems(3).DisplayValue(36)=   "/xDX1mNRKVoQAMa2pf///////////////wD////////////////////////39++caUKMUSFrKABz"
      Columns(2).ValueItems(3).DisplayValue(37)=   "MAB7IABaWTkA9/8A//8A9/8A9/8A//8xnoyEEABzKAB7KABzKABzMAB7IAAA7/cA//8A9/8A9/8A"
      Columns(2).ValueItems(3).DisplayValue(38)=   "//8Q395jCACMaVr///////////////8A////////////////////////////vaaElGE5aygAczAA"
      Columns(2).ValueItems(3).DisplayValue(39)=   "eyAAMaaUAP//APf/APf/APf/APf/APf/exgAczAAczAAczAAeyAAUmFCAP//APf/APf/APf/APf/"
      Columns(2).ValueItems(3).DisplayValue(40)=   "AP//WkkpYyAI7+/n////////////AP///////////////////////////97PvaV5Wns4CGsoAHsg"
      Columns(2).ValueItems(3).DisplayValue(41)=   "ACmmlAD//wD3/wD3/wD3/wD3/wDv/3sgAHMwAHMwAHMwAHsoAFJhQgD//wD3/wD3/wD3/wD3/wD/"
      Columns(2).ValueItems(3).DisplayValue(42)=   "/1pJKWMYAO/n5////////////wD////////////////////////////39++teWOUYTlrKAB7IABS"
      Columns(2).ValueItems(3).DisplayValue(43)=   "YUIA//8A//8A9/8A9/8A//8prpyEGABzMABzMABzMABzMABzKAAA7/cA//8A9/8A9/8A//8A5+9r"
      Columns(2).ValueItems(3).DisplayValue(44)=   "GAB7STH39/f///////////8A////////////////////////////////tZZzpXlSczAAcygAeygA"
      Columns(2).ValueItems(3).DisplayValue(45)=   "QoZrAPf/AP//AP//GM/GcygAczAAczAAczAAczAAczAAeygAWlkxCOfnAP//AP//CN/ea0khYxAA"
      Columns(2).ValueItems(3).DisplayValue(46)=   "zr6t////////////////AP///////////////////////////////////4RJGIxZMWsoAHMwAHsY"
      Columns(2).ValueItems(3).DisplayValue(47)=   "AHMoAFJpSms4EIQYAHMoAHMwAHMwAHMwAHMwAHMwAGswAHsoAIwgAFJhQmNJIYwwAJxhOeff1v//"
      Columns(2).ValueItems(3).DisplayValue(48)=   "/////////////////wD///////////////////////////////////+caUqUYTlzMABzMABzMABz"
      Columns(2).ValueItems(3).DisplayValue(49)=   "MAB7IABzKABzMABzMABzMABzMABzMABzMABzMAB7QRBzMABjGACEOBDGnozn18b/////////////"
      Columns(2).ValueItems(3).DisplayValue(50)=   "//////////////8A////////////////////////////////////59fOjFEpe0EQczAAczAAczAA"
      Columns(2).ValueItems(3).DisplayValue(51)=   "czAAczAAczAAczAAczAAczAAczgAczgIezgIczgAnHFK59/W////////////////////////////"
      Columns(2).ValueItems(3).DisplayValue(52)=   "////////////AP////////////////////////////////////f396V5WoRJIWsoAHMwAHMwAHMw"
      Columns(2).ValueItems(3).DisplayValue(53)=   "AHMwAHMwAHM4CHM4AGswAHM4CJRhOb2mjPfv5///////////////////////////////////////"
      Columns(2).ValueItems(3).DisplayValue(54)=   "/////////wD///////////////////////////////////////+9poyleVpzMABrKABrKABzMAB7"
      Columns(2).ValueItems(3).DisplayValue(55)=   "OAhzOAhzMACESRitjnPWx73/////////////////////////////////////////////////////"
      Columns(2).ValueItems(3).DisplayValue(56)=   "//////8A////////////////////////////////////////7+/ntZZzrY5zjFkxhEkYczAAaygA"
      Columns(2).ValueItems(3).DisplayValue(57)=   "jFkxxrac9/f3////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(3).DisplayValue(58)=   "////AP//////////////////////////////////////////////97WWe5xpSoRRKaV5Y/fv7///"
      Columns(2).ValueItems(3).DisplayValue(59)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(3).DisplayValue(60)=   "/wD/////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(3).DisplayValue(61)=   "//////////////////////////////////////////////////////////////////////////8A"
      Columns(2).ValueItems(3).DisplayValue(62)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(3).DisplayValue(63)=   "////////////////////////////////////////////////////////////////////////AA=="
      Columns(2).ValueItems(3).DisplayValue.vt=   9
      Columns(2).ValueItems(3)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(4)._DefaultItem=   0
      Columns(2).ValueItems(4).Value=   "005"
      Columns(2).ValueItems(4).Value.vt=   8
      Columns(2).ValueItems(4).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(2).ValueItems(4).DisplayValue(0)=   "bHQAADYOAABCTTYOAAAAAAAANgAAACgAAAAlAAAAIAAAAAEAGAAAAAAAAA4AAAAAAAAAAAAAAAAA"
      Columns(2).ValueItems(4).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(4).DisplayValue(2)=   "///////////////////37+fWz8b37+f/////////////////////////////////////////////"
      Columns(2).ValueItems(4).DisplayValue(3)=   "//8A////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(4).DisplayValue(4)=   "////9+/nzr61lGE5a0EpazAAa0EpnIZz////////////////////////////////////////////"
      Columns(2).ValueItems(4).DisplayValue(5)=   "AP////////////////////////////////////////////////////////////////////f375yG"
      Columns(2).ValueItems(4).DisplayValue(6)=   "c2tBKWswAEogAGswAGswAGswAEogALWWc////////////////////////////////////////wD/"
      Columns(2).ValueItems(4).DisplayValue(7)=   "///////////////////////////////////////////////////////37+/Ovq1rQSlKIABSIABr"
      Columns(2).ValueItems(4).DisplayValue(8)=   "MABjMBBrMABrMABrMABrQSlrMABrQSn37+f///////////////////////////////////8A////"
      Columns(2).ValueItems(4).DisplayValue(9)=   "////////////////////////////////////////////1se1nIZzazAASiAAazAAazAAezgIazAA"
      Columns(2).ValueItems(4).DisplayValue(10)=   "azAAa0EpazAAazAAazAAazAAazAA1s/G////////////////////////////////////AP//////"
      Columns(2).ValueItems(4).DisplayValue(11)=   "//////////////////////////////fv562WhJRhOVIgAGswAFIgAGswAGswAGswAGswAGswAGsw"
      Columns(2).ValueItems(4).DisplayValue(12)=   "AGswAGswAGswAGswAGswAFIgAFpZMf///////////////////////////////////wD/////////"
      Columns(2).ValueItems(4).DisplayValue(13)=   "///////////////////Wz8aUYTlKIABrMABSIAB7OAhrMAB7OAhrMAB7OAhrMAB7OAhrMAB7OAhr"
      Columns(2).ValueItems(4).DisplayValue(14)=   "MABrMABrMAB7OAhrMAB7OAhSIAD37+f///////////////////////////////8A////////////"
      Columns(2).ValueItems(4).DisplayValue(15)=   "////////1s/Ga0EpUiAAYzAQMWlaKZaEQo5zKZaEQo5zMWlaQo5zKZaEMWlaKZaEQo5zKZaEQo5z"
      Columns(2).ValueItems(4).DisplayValue(16)=   "MWlaQo5zKZaEQo5zMWlaQo5zpY57////////////////////////////////AP//////////////"
      Columns(2).ValueItems(4).DisplayValue(17)=   "/+ff1mswAGswAGswACG2rSGelDGelDGelCG2rTGelDGelDGelCG2rTGelCG2rTGelDGelDGelDGe"
      Columns(2).ValueItems(4).DisplayValue(18)=   "lDGelCG2rTGelDGelDGelJxpSv///////////////////////////////wD///////////////+l"
      Columns(2).ValueItems(4).DisplayValue(19)=   "jntrMABrQSlrMAAA5/drMABSIABrMABrMABrMABSIABrMABrMABrMABrMABrMABrMAB7OAhSIABr"
      Columns(2).ValueItems(4).DisplayValue(20)=   "MABrMABrMABrMABrMABKIADOvq3///////////////////////////8A////////////////lHFa"
      Columns(2).ValueItems(4).DisplayValue(21)=   "SiAAazAAazAAGM/GMWlaazAAazAAezgIazAAazAAa0EpezgIazAAazAAazAAezgIa0EpazAAazAA"
      Columns(2).ValueItems(4).DisplayValue(22)=   "ezgIazAAazAAa0EpazAAlHFa////////////////////////////AP///////////////6WOe2tB"
      Columns(2).ValueItems(4).DisplayValue(23)=   "GFIgAGswADFpWhjPxlIgAGswAGtBKWswAGswAGswAGMwEGswAGswAGswAGMwEGswAGswAGswAGtB"
      Columns(2).ValueItems(4).DisplayValue(24)=   "KWswAGswAGswAFIgAGswAPfv7////////////////////////wD////////////////n39aUcVqE"
      Columns(2).ValueItems(4).DisplayValue(25)=   "SSFrMAB7OAgA9/9rQRhrMACESSFrMABrMABrMABrMABrMAB7OAhrMABrMABrMAB7OAhrMAB7OAhr"
      Columns(2).ValueItems(4).DisplayValue(26)=   "MABrMABrMABrMABKIADWx7X///////////////////////8A////////////////////lGE5c1lC"
      Columns(2).ValueItems(4).DisplayValue(27)=   "azAAUiAAEM/OMWlaazAAa0EpazAAa0EpOWFKa0EpWlkxUjgYazAAMWlaazAAUiAAWlkxUjgYazAA"
      Columns(2).ValueItems(4).DisplayValue(28)=   "OWFKazAAazAAa0EYc1lC////////////////////////AP///////////////////62WhJRhOWsw"
      Columns(2).ValueItems(4).DisplayValue(29)=   "AHs4CDFpWiG2rWswAHs4CFI4GAjf5wD3/yG2rQDn9yG2rWtBKQD3/zFpWns4CBDPzlKWhFI4GAD3"
      Columns(2).ValueItems(4).DisplayValue(30)=   "/1pZMWswABDPzhDHxs6+rf///////////////////wD////////////////////Wz8aEUSGESSFS"
      Columns(2).ValueItems(4).DisplayValue(31)=   "IABaWTEQx8ZrMABrQSlrMAAQz85CjnMxaVoA9/8A5/drMAAA5/cxnpRSIAAhtq0Qx8ZrMAAA5/dK"
      Columns(2).ValueItems(4).DisplayValue(32)=   "jnM5YUoA9/9SOBichnP///////////////////8A////////////////////////lGE5SnFSazAA"
      Columns(2).ValueItems(4).DisplayValue(33)=   "SiAAIbatUjgYazAAazAAGM/GEM/OezgIMZ6UAPf/UiAAIbatAPf/a0EYMWlaAPf/UiAAIbatEM/O"
      Columns(2).ValueItems(4).DisplayValue(34)=   "a0EYAPf/WlkxazAA9+/n////////////////AP///////////////////////9bPxoRJIWtBKWsw"
      Columns(2).ValueItems(4).DisplayValue(35)=   "ACG2rVpZMWswAGswACmWhAD3/1IgADlhSgDn92tBGDGelAD3/yG2rTGelAD3/2swADFpWgD3/1Ig"
      Columns(2).ValueItems(4).DisplayValue(36)=   "ACG2rQDn9zlhSs6+tf///////////////wD///////////////////////////9aWTGUYTlrMABC"
      Columns(2).ValueItems(4).DisplayValue(37)=   "jnMxnpR7OAhrMABaWTEQz85aWTFSIAAYz8Y5YUpaWTEQz85aWTEQz84Q195SIABaWTEQz857OAhS"
      Columns(2).ValueItems(4).DisplayValue(38)=   "OBghtq0Qz86tloT///////////////8A////////////////////////////nIZzhGFKazAAMWla"
      Columns(2).ValueItems(4).DisplayValue(39)=   "GM/GUiAAazAAUiAAazAAa0EpOWFKazAAazAAa0EpazAAUiAAezgIUiAAazAAUiAAQo5zc1lCazAA"
      Columns(2).ValueItems(4).DisplayValue(40)=   "UiAAazAAUiAA9+/n////////////AP///////////////////////////9bPxpyGc2swAGswAADn"
      Columns(2).ValueItems(4).DisplayValue(41)=   "92tBGGswAGswAGswAFpZMRDPziG2rWswAGswAGswAGswAGswAHs4CGswAGtBGBDPzkqOc2swAHs4"
      Columns(2).ValueItems(4).DisplayValue(42)=   "CGswAGswANbPxv///////////wD////////////////////////////37+eUcVqUYTlSIAAhtq0x"
      Columns(2).ValueItems(4).DisplayValue(43)=   "npR7OAhSIABrMABSIAB7OAhrQSlrMABSIAB7OAhSIABrMABSIABrMABSIABrMABSIABrMABSIABr"
      Columns(2).ValueItems(4).DisplayValue(44)=   "MABrQSn37+f///////////8A////////////////////////////////tZZzSnFSazAAUiAAAO//"
      Columns(2).ValueItems(4).DisplayValue(45)=   "AOf3CN/nAOf3GM/GAOf3ENfeEM/OENfeAOf3CN/nEM/OAO//AOf3CN/nAOf3ENfeAOf3CN/nMWla"
      Columns(2).ValueItems(4).DisplayValue(46)=   "zr6t////////////////AP////////////////////////////////fv72swAIRRIWswAFIgAGsw"
      Columns(2).ValueItems(4).DisplayValue(47)=   "AFI4GGtBGGtBKWswAGtBKWswAGtBKWswAFI4GGtBGGtBKWswAGtBKWswAFI4GFpZMVKWhOff1v//"
      Columns(2).ValueItems(4).DisplayValue(48)=   "/////////////////wD///////////////////////////////////9aWTGUYTlrMACESSFrMABr"
      Columns(2).ValueItems(4).DisplayValue(49)=   "MABrMABrMABrMAB7OAhrMABrMABrMABrMABrMACESSFrMABrMABrQSm1lnPWz8b/////////////"
      Columns(2).ValueItems(4).DisplayValue(50)=   "//////////////8A////////////////////////////////////9+/nhGFKazAASiAAazAAUiAA"
      Columns(2).ValueItems(4).DisplayValue(51)=   "azAAazAAazAAUiAAazAAazAAazAAUiAAazAAa0EplGE55+fe////////////////////////////"
      Columns(2).ValueItems(4).DisplayValue(52)=   "////////////AP////////////////////////////////////f375yGc1pZMXs4CGswAGswAGsw"
      Columns(2).ValueItems(4).DisplayValue(53)=   "AHs4CGswAIRJIWswAGswAGtBKZRhOa2WhP//////////////////////////////////////////"
      Columns(2).ValueItems(4).DisplayValue(54)=   "/////////wD///////////////////////////////////////+ljnuUYTlrQSlrMABSIABrMABr"
      Columns(2).ValueItems(4).DisplayValue(55)=   "QSlrMABrMABrMACljnvWz8b/////////////////////////////////////////////////////"
      Columns(2).ValueItems(4).DisplayValue(56)=   "//////8A////////////////////////////////////////9+/nrZaEtZZza0EphEkha0EpazAA"
      Columns(2).ValueItems(4).DisplayValue(57)=   "a0Epzr6t9/fv////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(4).DisplayValue(58)=   "////AP////////////////////////////////////////////fv58amlJRhOYRRIbWWc/fv7///"
      Columns(2).ValueItems(4).DisplayValue(59)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(4).DisplayValue(60)=   "/wD////////////////////////////////////////////////////39+//////////////////"
      Columns(2).ValueItems(4).DisplayValue(61)=   "//////////////////////////////////////////////////////////////////////////8A"
      Columns(2).ValueItems(4).DisplayValue(62)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(4).DisplayValue(63)=   "////////////////////////////////////////////////////////////////////////AA=="
      Columns(2).ValueItems(4).DisplayValue.vt=   9
      Columns(2).ValueItems(4)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(5)._DefaultItem=   0
      Columns(2).ValueItems(5).Value=   "006"
      Columns(2).ValueItems(5).Value.vt=   8
      Columns(2).ValueItems(5).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(2).ValueItems(5).DisplayValue(0)=   "bHQAADYOAABCTTYOAAAAAAAANgAAACgAAAAlAAAAIAAAAAEAGAAAAAAAAA4AAAAAAAAAAAAAAAAA"
      Columns(2).ValueItems(5).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(5).DisplayValue(2)=   "///////////////////v5+e97+/v5+f/////////////////////////////////////////////"
      Columns(2).ValueItems(5).DisplayValue(3)=   "//8A////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(5).DisplayValue(4)=   "////////xqaUlGE5WjAQaygAa0EhnIZr////////////////////////////////////////////"
      Columns(2).ValueItems(5).DisplayValue(5)=   "AP///////////////////////////////////////////////////////////////////9b//5yG"
      Columns(2).ValueItems(5).DisplayValue(6)=   "a2tBIWsoAEogAGsoAGsoAGsoAEogALWWc////////////////////////////////////////wD/"
      Columns(2).ValueItems(5).DisplayValue(7)=   "///////////////////////////////////////////////////////37+/Ovq1rQSFKIABSIABr"
      Columns(2).ValueItems(5).DisplayValue(8)=   "KABSIABrKABSIABrKABrQSFrKABrQSHv5+f///////////////////////////////////8A////"
      Columns(2).ValueItems(5).DisplayValue(9)=   "////////////////////////////////////////////1se1nIZraygASiAAaygAaygAczgIaygA"
      Columns(2).ValueItems(5).DisplayValue(10)=   "aygAa0EhaygAaygAaygAaygAaygA1s/G////////////////////////////////////AP//////"
      Columns(2).ValueItems(5).DisplayValue(11)=   "/////////////////////////////+/n56WOe5RhOVIgAGsoAFIgAGsoAGsoAGsoAFIgAGsoAGso"
      Columns(2).ValueItems(5).DisplayValue(12)=   "AGsoAFIgAGsoAGsoAGsoAFIgAGNZMf///////////////////////////////////wD/////////"
      Columns(2).ValueItems(5).DisplayValue(13)=   "///////////////////Wz8aUYTlKOBhrKABrKABzOAhrKABrKABrKABzOAhrKABrKABrKABrKABr"
      Columns(2).ValueItems(5).DisplayValue(14)=   "KABrKABrKABzOAhKIABrKABKIADv5+f///////////////////////////////8A////////////"
      Columns(2).ValueItems(5).DisplayValue(15)=   "////////zr6ta0EhUiAAEL7GENfWEL7GMZ6UMZ6UOY5zMXlrWoZrQllCa0EhUiAAaygAUiAAaygA"
      Columns(2).ValueItems(5).DisplayValue(16)=   "UiAAaygAWjAQaygAaygAaygAlHFa////////////////////////////////AP//////////////"
      Columns(2).ValueItems(5).DisplayValue(17)=   "/+ff1msoAGsoAEo4GBjf50ogAFqGaxiurTmOczF5a1qGaxiurSmupRDX1jnPzhDX1imupTmOc2NZ"
      Columns(2).ValueItems(5).DisplayValue(18)=   "MWsoAGsoAGsoAGsoAEogAJRhOf///////////////////////////////wD///////////////+l"
      Columns(2).ValueItems(5).DisplayValue(19)=   "jntrKABrKABrQRgQ19ZrKABrKABrKABrQSE5jnMxnpQ5jnMxeWtjWTFrQSFjWTExeWtavrUQvsYQ"
      Columns(2).ValueItems(5).DisplayValue(20)=   "19YA7/c5z84prq1ahmtKOBjWz8b///////////////////////////8A////////////////lGE5"
      Columns(2).ValueItems(5).DisplayValue(21)=   "a0EhaygAQllCMZ6USiAAaygAaygAaygAUiAAaygAUiAAa0EYQllCc2lKMXlrY1kxMXlrY1kxQllC"
      Columns(2).ValueItems(5).DisplayValue(22)=   "lGE5MXlrOY5zMZ6UOc/OKa6t////////////////////////////AP///////////////6WOe2tB"
      Columns(2).ValueItems(5).DisplayValue(23)=   "GFIgADmOczF5a2soAGsoAGsoAGtBIWsoAFIgAGsoAFIgAGsoAFIgAGsoAFowEGsoAEo4GGsoAGtB"
      Columns(2).ValueItems(5).DisplayValue(24)=   "IWNZMUJZQmNZMWtBIUogAL3v7////////////////////////wD////////////////Wx7WUcVpz"
      Columns(2).ValueItems(5).DisplayValue(25)=   "OAgYrq05jnNrKAA5jnMYrq1ahmsxeWtahmtKOBhjWTFrKABzOAhKIABzOAhrKABzOAhrKABzOAhr"
      Columns(2).ValueItems(5).DisplayValue(26)=   "KABrKABrKABzOAg5jnO97+////////////////////////8A////////////////////lGE5jFEh"
      Columns(2).ValueItems(5).DisplayValue(27)=   "OY5zMXlraygAEL7GWoZrOc/OSnFSKa6tENfWc2lKOY5za0EhMZ6Uc2lKa0EYSjgYaygASjgYaygA"
      Columns(2).ValueItems(5).DisplayValue(28)=   "UiAAaygAUiAAOY5zMZ6U////////////////////////AP///////////////////62WhJxxSjGe"
      Columns(2).ValueItems(5).DisplayValue(29)=   "lGNZMWsoAGNZMUo4GFqGa0JZQimupRDX1lqGawDv92tBGADv9znPzkogACmupTGelDGelDF5a1qG"
      Columns(2).ValueItems(5).DisplayValue(30)=   "azF5a3M4CBC+xkJZQtbPxv///////////////////wD////////////////////Wz8ZzaUo5z85r"
      Columns(2).ValueItems(5).DisplayValue(31)=   "QSFrKABSIABrKABSIABrKABrQSFrKABrQSFrKABrQSE5jnMxeWtzrpwA7/c5z84prq0Q19YA9/9j"
      Columns(2).ValueItems(5).DisplayValue(32)=   "WTFSIAAprqVKOBichmv///////////////////8A////////////////////////lGE5ENfWa0EY"
      Columns(2).ValueItems(5).DisplayValue(33)=   "SiAAaygAaygAaygAMXlrKa6lSiAAOY5zGK6tY1kxSjgYaygASiAAczgIUiAAY1kxaygAWoZra0Eh"
      Columns(2).ValueItems(5).DisplayValue(34)=   "a0EYGK6taygAa0Eh7+fn////////////////AP///////////////////////73v7xDX1mtBIWso"
      Columns(2).ValueItems(5).DisplayValue(35)=   "AFIgAGsoAFIgAHNpShjf52NZMTGelCmupQDv9zmOcxDX1jGelGtBIWsoAGsoAGsoAGsoAGsoAEo4"
      Columns(2).ValueItems(5).DisplayValue(36)=   "GCmupWtBIWsoAM6+tf///////////////wD///////////////////////////85jnOUYTkxeWtr"
      Columns(2).ValueItems(5).DisplayValue(37)=   "QRhrQSFjWTFrQSFjWTFrKABjWTFKOBgprqUxeWsprqUA7/djWTFrKABzOAhrKABrKABrKABjWTEY"
      Columns(2).ValueItems(5).DisplayValue(38)=   "rq1rKABrKACUYTn///////////////8A////////////////////////////tZZzc2lKY1kxMXlr"
      Columns(2).ValueItems(5).DisplayValue(39)=   "Y1kxMXlrQllCc2lKQllCSjgYaygASjgYaygAWjAQaygAUiAAaygAUiAAaygAaygAaygAMXlrOY5z"
      Columns(2).ValueItems(5).DisplayValue(40)=   "UiAAaygAWjAQ7+fn////////////AP///////////////////////////9b//1rn5xDX1imupRiu"
      Columns(2).ValueItems(5).DisplayValue(41)=   "rTmOczmOc1qGazF5a2NZMTF5a1qGazF5a1qGa0JZQmNZMTF5a3NpSmtBIXM4CGsoADmOczmOc3M4"
      Columns(2).ValueItems(5).DisplayValue(42)=   "CEogAGsoANbPxv///////////wD////////////////////////////v5+eUcVqUYTlKOBhjWTFz"
      Columns(2).ValueItems(5).DisplayValue(43)=   "aUo5jnMprq0prqUQ19YprqUprq0xnpQxeWtahmsxeWtrKABrQSFrKAAxeWtrKAAxnpRCWUJSIABr"
      Columns(2).ValueItems(5).DisplayValue(44)=   "KABrQSHv5+f///////////8A////////////////////////////////tZZzlHFaaygAaygAaygA"
      Columns(2).ValueItems(5).DisplayValue(45)=   "aygAczgISiAAaygASjgYa0EYMXlrOY5zENfWOc/OAO/3GN/nOc/OWoZra0EhKa6lQllCaygAaygA"
      Columns(2).ValueItems(5).DisplayValue(46)=   "zr6t////////////////AP////////////////////////////////fv72soAIxRIWsoAGtBIWso"
      Columns(2).ValueItems(5).DisplayValue(47)=   "AGsoAGsoAFIgAGsoAGsoAGsoAFIgAGsoAFIgAHM4CGtBIWNZMTGelDGelCmurWNZMZRxWuff1v//"
      Columns(2).ValueItems(5).DisplayValue(48)=   "/////////////////wD///////////////////////////////////9jWTGUYTlrKABzOAhrKABr"
      Columns(2).ValueItems(5).DisplayValue(49)=   "KABrKABrKABrKABzOAhrKABrKABrKABrKABKIACESSFrKABrKABrQSG1lnPWz8b/////////////"
      Columns(2).ValueItems(5).DisplayValue(50)=   "//////////////8A////////////////////////////////////7+fnhGFCaygAUiAAaygAUiAA"
      Columns(2).ValueItems(5).DisplayValue(51)=   "aygAaygAaygAUiAAaygAaygAaygAUiAAaygAa0EhlGE5ve/v////////////////////////////"
      Columns(2).ValueItems(5).DisplayValue(52)=   "////////////AP///////////////////////////////////9b//5yGa2NZMXM4CGsoAGsoAGso"
      Columns(2).ValueItems(5).DisplayValue(53)=   "AHM4CGsoAGtBGGsoAGsoAGtBIZRhOa2WhP//////////////////////////////////////////"
      Columns(2).ValueItems(5).DisplayValue(54)=   "/////////wD///////////////////////////////////////+tloSUYTlrQSFrKABSIABrKABr"
      Columns(2).ValueItems(5).DisplayValue(55)=   "QSFrKABrKABrKACljnvWz8b/////////////////////////////////////////////////////"
      Columns(2).ValueItems(5).DisplayValue(56)=   "//////8A////////////////////////////////////////7+fnrZaEtZZza0EhhEkhSiAAaygA"
      Columns(2).ValueItems(5).DisplayValue(57)=   "a0Ehzr6t1v//////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(5).DisplayValue(58)=   "////AP///////////////////////////////////////////+/n58amlJRhOYxRIbWWc/fv7///"
      Columns(2).ValueItems(5).DisplayValue(59)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(5).DisplayValue(60)=   "/wD////////////////////////////////////////////////////W////////////////////"
      Columns(2).ValueItems(5).DisplayValue(61)=   "//////////////////////////////////////////////////////////////////////////8A"
      Columns(2).ValueItems(5).DisplayValue(62)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(5).DisplayValue(63)=   "////////////////////////////////////////////////////////////////////////AA=="
      Columns(2).ValueItems(5).DisplayValue.vt=   9
      Columns(2).ValueItems(5)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(6)._DefaultItem=   0
      Columns(2).ValueItems(6).Value=   "007"
      Columns(2).ValueItems(6).Value.vt=   8
      Columns(2).ValueItems(6).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(2).ValueItems(6).DisplayValue(0)=   "bHQAADYOAABCTTYOAAAAAAAANgAAACgAAAAlAAAAIAAAAAEAGAAAAAAAAA4AAAAAAAAAAAAAAAAA"
      Columns(2).ValueItems(6).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(6).DisplayValue(2)=   "///////////////////v7+/e39bv7+f/////////////////////////////////////////////"
      Columns(2).ValueItems(6).DisplayValue(3)=   "//8A////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(6).DisplayValue(4)=   "////////xr6tjGlSYzAYWiAAWigQpY57////////////////////////////////////////////"
      Columns(2).ValueItems(6).DisplayValue(5)=   "AP////////////////////////////////////////////////////////////////////f375R5"
      Columns(2).ValueItems(6).DisplayValue(6)=   "a2MwKVogAGMgAGsoAHMoAGsoAFIYAKWOhP///////////////////////////////////////wD/"
      Columns(2).ValueItems(6).DisplayValue(7)=   "///////////////////////////////////////////////////////39/etx8Z7SSlSCABaGABr"
      Columns(2).ValueItems(6).DisplayValue(8)=   "KABzMABzMABzMABzMABzMABrKABjMBj39/f///////////////////////////////////8A////"
      Columns(2).ValueItems(6).DisplayValue(9)=   "////////////////////////////////////////////zse9c8/OCOfvAP//MXFSa0EQczAAczAA"
      Columns(2).ValueItems(6).DisplayValue(10)=   "czAAczAAczAAczAAczAAczAAShAAzr61////////////////////////////////////AP//////"
      Columns(2).ValueItems(6).DisplayValue(11)=   "/////////////////////////////+/v57WmlHtRQlogAEJJMTGWhBDXzgD//yG2rUKGa3sgAHMw"
      Columns(2).ValueItems(6).DisplayValue(12)=   "AHMwAHMwAHMwAHMwAHMwAFogAGtBKf///////////////////////////////////wD/////////"
      Columns(2).ValueItems(6).DisplayValue(13)=   "///////////////////e19aEaVJaIABaGABrIABzMAB7KAB7IABzKABjSSFCeWMA9/9ChnN7IABz"
      Columns(2).ValueItems(6).DisplayValue(14)=   "MABzMABzMABzMABzMABzMABKEAD39/f///////////////////////////////8A////////////"
      Columns(2).ValueItems(6).DisplayValue(15)=   "////////3tfGa0EhQgAAYxgAczAAczAAczAAczAAczAAczAAczAAcygAjBAASnFSGL69UmFCczAA"
      Columns(2).ValueItems(6).DisplayValue(16)=   "czAAczAAczAAczAAczAAWhgAnIZz////////////////////////////////AP//////////////"
      Columns(2).ValueItems(6).DisplayValue(17)=   "/97XzmMwEGMgAHMwAHMwAHMwAHMwAHMwAHMwAHMwAHMwAHMwAHMwAHsgAEKGa0pxUkpxWkppSmNJ"
      Columns(2).ValueItems(6).DisplayValue(18)=   "IXsoAHMwAHMwAHMwAGsoAHNJMff39////////////////////////////wD///////////////+l"
      Columns(2).ValueItems(6).DisplayValue(19)=   "hnNjGABzMABzMABzMABzMABzMABzMABzMABzMABzMABzMABzMAB7IAA5jnM5lns5lns5loRCeWNj"
      Columns(2).ValueItems(6).DisplayValue(20)=   "SSFzKABzMABzMABzMABaIADOvrX///////////////////////////8A////////////////nGlK"
      Columns(2).ValueItems(6).DisplayValue(21)=   "ezgIczAAczAAczAAczAAazgYazAIeygAeygAcygAczAAczAAeyAAOZZ7OYZrAPf/MZaEOY5zQoZr"
      Columns(2).ValueItems(6).DisplayValue(22)=   "azgYcygAczAAczAAWhgAlHlr////////////////////////////AP///////////////62Oa4xZ"
      Columns(2).ValueItems(6).DisplayValue(23)=   "MXMwAHMwAHMwAHMoABjHvVJpSkpxWlpZOWtBEHsoAHMoAIQYADGejEKGawjv7wjv70KGazmOczGW"
      Columns(2).ValueItems(6).DisplayValue(24)=   "hHMwAHMoAHMwAGsoAFooCPfv7////////////////////////wD////////////////Wvq2laUKE"
      Columns(2).ValueItems(6).DisplayValue(25)=   "MABzKAB7IABSaUoI395ChmsxnoxCeWMI1945hmtjSSFzKAApppRChmsA7+8A//8hvrVChms5jnMp"
      Columns(2).ValueItems(6).DisplayValue(26)=   "rqV7KABzKABzMABKCADGtq3///////////////////////8A////////////////////UratKb61"
      Columns(2).ValueItems(6).DisplayValue(27)=   "UllCa0Epa0EQY1EpSnFaMZ6EGMfGSnFSGMfGEM/OczAYKaacQoZrAO/3APf/AP//MZaEOY57MZ6E"
      Columns(2).ValueItems(6).DisplayValue(28)=   "Ka6leygAczAAYyAAa0Eh////////////////////////AP///////////////////86ehHuGazGW"
      Columns(2).ValueItems(6).DisplayValue(29)=   "hCmmlDGehEJ5a1pRMXMwAFpROWs4CCGupSmmnIwIACG2rUKGawDv9wD3/wD3/wD3/0pxWjGWhBjH"
      Columns(2).ValueItems(6).DisplayValue(30)=   "vTGejIQYAHMwAFIYAOff3v///////////////////wD////////////////////n39acYTGMOBBz"
      Columns(2).ValueItems(6).DisplayValue(31)=   "KABjSSlKcVIpppQI194xppxzKAB7KABrOAiMAAAYvr1ChmsA9/8A9/8A9/8A//8I5+daWTE5jnsA"
      Columns(2).ValueItems(6).DisplayValue(32)=   "9/9SaUqEGABSGACtnoz///////////////////8A////////////////////////lGFCnGlKczAA"
      Columns(2).ValueItems(6).DisplayValue(33)=   "cygAeygAhBgAhBAAWlk5EM/OIa6lSnFSY1EhEM/OSnFaAPf/APf/APf/APf/AP//GMfGQnljSnFS"
      Columns(2).ValueItems(6).DisplayValue(34)=   "AP//SnFacyAAYzAY9/f3////////////////AP///////////////////////97HvYRJIXtBEHMw"
      Columns(2).ValueItems(6).DisplayValue(35)=   "AHMwAHMwAHMwAHsgAHsYAEKGa0J5Y0pxWlpROUpxUgD//wD3/wD3/wD3/wD3/wD3/zmWhCmmnDGe"
      Columns(2).ValueItems(6).DisplayValue(36)=   "hAD//2tBGFoYAMa2pf///////////////wD////////////////////////39++caUKMUSFrKABz"
      Columns(2).ValueItems(6).DisplayValue(37)=   "MABzMABzMABzMAB7KABKaUoA7/cYz8Y5jnsQ19YA//8A9/8A9/8A9/8A9/8A//8Q19ZKcVoYvr1K"
      Columns(2).ValueItems(6).DisplayValue(38)=   "cVJzOAhaGACMaWP///////////////8A////////////////////////////vaaElGE5aygAczAA"
      Columns(2).ValueItems(6).DisplayValue(39)=   "czAAczAAczAAczAAeyAAOYZzAP//AP//AP//CN/nAOfvAPf/APf/AP//GL69OYZrMZaEY0kheyAA"
      Columns(2).ValueItems(6).DisplayValue(40)=   "czAAaygAYzAQ7+fn////////////AP///////////////////////////97PvaV5Wns4CHMoAHMw"
      Columns(2).ValueItems(6).DisplayValue(41)=   "AHMwAHMwAHMwAHMwAHsoADmWhAD//xDPziG+rUKGawD//wD39ymunDmOczmWe2s4AHsgAHMwAHMw"
      Columns(2).ValueItems(6).DisplayValue(42)=   "AHMoAFooAO/n5////////////wD////////////////////////////39++teWOUYTlrKABzMABz"
      Columns(2).ValueItems(6).DisplayValue(43)=   "MABzMABzMABzMABzKAB7GAAhtqUQz84hvrUxnowI3+cxloQpppQxppRaUSlzKABzMABzMABzMABj"
      Columns(2).ValueItems(6).DisplayValue(44)=   "IAB7STn39/f///////////8A////////////////////////////////tZZzpXlSczAAczAAczAA"
      Columns(2).ValueItems(6).DisplayValue(45)=   "czAAczAAczAAczAAcygAcygAENfWGM/GCO/vMZaEGL69OZZ7QoZjcygAczAAczAAczAAczAAYyAA"
      Columns(2).ValueItems(6).DisplayValue(46)=   "zr6t////////////////AP///////////////////////////////////4RJGIxZMWsoAHMwAHMw"
      Columns(2).ValueItems(6).DisplayValue(47)=   "AHMwAHMwAHMwAHMwAHsoAGNBGBDX1kJ5Yxi+vVJhOSmelIQgAHs4CHMwAHMwAIRJIZRpSuff1v//"
      Columns(2).ValueItems(6).DisplayValue(48)=   "/////////////////wD///////////////////////////////////+caUqUYTlzMABzMABzMABz"
      Columns(2).ValueItems(6).DisplayValue(49)=   "MABzMABzMABzMABzMABzMABrOBBrOBBaUSlzKAB7QRBzMABjGAB7SRjGppTe187/////////////"
      Columns(2).ValueItems(6).DisplayValue(50)=   "//////////////8A////////////////////////////////////59fOjFEpe0EQczAAczAAczAA"
      Columns(2).ValueItems(6).DisplayValue(51)=   "czAAczAAczAAczAAczAAczAAczAAezAAezgIczAAnHFK59/W////////////////////////////"
      Columns(2).ValueItems(6).DisplayValue(52)=   "////////////AP////////////////////////////////////f396V5WoRJIWsoAHMwAHMwAHMw"
      Columns(2).ValueItems(6).DisplayValue(53)=   "AHMwAHMwAHM4CHM4AHMwAHM4CJRhOb2mjPfv5///////////////////////////////////////"
      Columns(2).ValueItems(6).DisplayValue(54)=   "/////////wD///////////////////////////////////////+9poyleVpzMABrKABrKABzMAB7"
      Columns(2).ValueItems(6).DisplayValue(55)=   "OAhzOAhzMACESRitjnPWx73/////////////////////////////////////////////////////"
      Columns(2).ValueItems(6).DisplayValue(56)=   "//////8A////////////////////////////////////////7+/ntZZztY5zlFkxhEkYczAAaygA"
      Columns(2).ValueItems(6).DisplayValue(57)=   "jFkxxrac9/f3////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(6).DisplayValue(58)=   "////AP/////////////////////////////////////////////397WWe5xpSoRRKaV5Y/fv7///"
      Columns(2).ValueItems(6).DisplayValue(59)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(6).DisplayValue(60)=   "/wD/////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(6).DisplayValue(61)=   "//////////////////////////////////////////////////////////////////////////8A"
      Columns(2).ValueItems(6).DisplayValue(62)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(6).DisplayValue(63)=   "////////////////////////////////////////////////////////////////////////AA=="
      Columns(2).ValueItems(6).DisplayValue.vt=   9
      Columns(2).ValueItems(6)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(7)._DefaultItem=   0
      Columns(2).ValueItems(7).Value=   "008"
      Columns(2).ValueItems(7).Value.vt=   8
      Columns(2).ValueItems(7).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(2).ValueItems(7).DisplayValue(0)=   "bHQAADYOAABCTTYOAAAAAAAANgAAACgAAAAlAAAAIAAAAAEAGAAAAAAAAA4AAAAAAAAAAAAAAAAA"
      Columns(2).ValueItems(7).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(7).DisplayValue(2)=   "///////////////////37+fWz8b37+f/////////////////////////////////////////////"
      Columns(2).ValueItems(7).DisplayValue(3)=   "//8A////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(7).DisplayValue(4)=   "////9+/nzr61lGE5a0EpazAAa0EpnIZz////////////////////////////////////////////"
      Columns(2).ValueItems(7).DisplayValue(5)=   "AP////////////////////////////////////////////////////////////////////f375yG"
      Columns(2).ValueItems(7).DisplayValue(6)=   "c2tBKWswAEogAGswAGswAGswAEogALWWc////////////////////////////////////////wD/"
      Columns(2).ValueItems(7).DisplayValue(7)=   "///////////////////////////////////////////////////////37+/Ovq1rQSlKIABSIABr"
      Columns(2).ValueItems(7).DisplayValue(8)=   "MABjMBBrMABrMABrMABrQSlrMABrQSn37+f///////////////////////////////////8A////"
      Columns(2).ValueItems(7).DisplayValue(9)=   "////////////////////////////////////////////1se1nIZzazAASiAAazAAazAAezgIazAA"
      Columns(2).ValueItems(7).DisplayValue(10)=   "azAAazAAazAAazAAazAAazAAazAA1s/G////////////////////////////////////AP//////"
      Columns(2).ValueItems(7).DisplayValue(11)=   "//////////////////////////////fv562WhJRhOVIgAGswAFIgAGswAGswAGswAFIgAFpZMVpZ"
      Columns(2).ValueItems(7).DisplayValue(12)=   "MWtBGFIgAGswAGswAGswAFIgAFpZMf///////////////////////////////////wD/////////"
      Columns(2).ValueItems(7).DisplayValue(13)=   "///////////////////Wz8aUcVpKIABrMABrMAB7OAhrMABrMABrMABrMABKeVpChmtrMABaWTFr"
      Columns(2).ValueItems(7).DisplayValue(14)=   "MABrQRhrMAB7OAhrMABrMABrMAD37+f///////////////////////////////8A////////////"
      Columns(2).ValueItems(7).DisplayValue(15)=   "////////1s/Ga0EpSiAAUiAAazAAazAAazAAYzAQazAAazAAWlkxMZ6UWlkxa0EpazAAa0EpWlkx"
      Columns(2).ValueItems(7).DisplayValue(16)=   "SnlaWlkxMZ6UazAAazAAazAAlHFa////////////////////////////////AP//////////////"
      Columns(2).ValueItems(7).DisplayValue(17)=   "/+ff1mswAGswAGswAGswAGswAGswAGswAGswAGswADGelEp5WmswAGswAGswAEp5WimupVpZMXs4"
      Columns(2).ValueItems(7).DisplayValue(18)=   "CDGelAD3/0p5WmswAGswAJRhOf///////////////////////////////wD///////////////+l"
      Columns(2).ValueItems(7).DisplayValue(19)=   "jntrMABrQSlrMABrMABrMABSIABrMABrMABSYUI5jnt7OAhrMABrMABKeVpaWTFrMABaWTFKeVop"
      Columns(2).ValueItems(7).DisplayValue(20)=   "rqUA9/8prqVSIABrMABKIADOvq3///////////////////////////8A////////////////lHFa"
      Columns(2).ValueItems(7).DisplayValue(21)=   "SiAAazAAazAAazAAa0EpazAAazAAQoZrKa6cWlkxYzAQazAAUmFCUmlKazAAezgISiAAOY5zENfW"
      Columns(2).ValueItems(7).DisplayValue(22)=   "APf/ENfWazAAa0EpazAAlHFa////////////////////////////AP///////////////6WOe2tB"
      Columns(2).ValueItems(7).DisplayValue(23)=   "GFIgAGswAFIgAGswAGswAGswAEpxUlJhQmMwEFJhQjmOeymupWtBKWswAFIgAFpZMUp5WhjPxgDn"
      Columns(2).ValueItems(7).DisplayValue(24)=   "9wjf52tBKVJpSlIgAGswAPfv7////////////////////////wD////////////////n39aUcVqE"
      Columns(2).ValueItems(7).DisplayValue(25)=   "SSFrMABrMABrMAB7OAhrMABChmtrMAB7OAhrMAA5jnNSYUJaWTFKeVprMAAQ19YA9/8A9/8A9/8A"
      Columns(2).ValueItems(7).DisplayValue(26)=   "5/daWTExnpR7OAhKIADn39b///////////////////////8A////////////////////lGE5e1lC"
      Columns(2).ValueItems(7).DisplayValue(27)=   "azAAazAAazAAUiAAazAASnlaazAAUiAASiAAMZ6UazAAUiAAazAAIb61APf/AOf3APf/AOf3APf/"
      Columns(2).ValueItems(7).DisplayValue(28)=   "OY57OY5zazAAazAAa0Ep////////////////////////AP///////////////////62WhJRhOWsw"
      Columns(2).ValueItems(7).DisplayValue(29)=   "AHs4CGswAGswAGtBKVpZMUogAHs4CDGelFpZMVIgAGswAFJhQgDv/wDn9wD3/wD3/wDv/wD3/ymu"
      Columns(2).ValueItems(7).DisplayValue(30)=   "pRDX1ns4CGswAGswANbPxv///////////////////wD////////////////////Wz8aEUSGESSFS"
      Columns(2).ValueItems(7).DisplayValue(31)=   "IABrMABSIAB7OAh7WUJaWTEhvrUA9/8A5/cQ19YxnpQprqUA5/cA9/8A5/cA9/8A5/cA9/8xnpQp"
      Columns(2).ValueItems(7).DisplayValue(32)=   "rqVSIABrMABSIACchnP///////////////////8A////////////////////////lGE5SnFSazAA"
      Columns(2).ValueItems(7).DisplayValue(33)=   "SiAAQoZrAPf/APf/APf/Ka6la0Epa0EYKa6ca0EYUmFCKa6lAPf/APf/APf/APf/APf/OY5zKa6c"
      Columns(2).ValueItems(7).DisplayValue(34)=   "azAAazAAazAAa0Ep9+/n////////////////AP///////////////////////9bPxoRJIWtBKWsw"
      Columns(2).ValueItems(7).DisplayValue(35)=   "ADmOewD3/wDn9wD3/zGelFpZMWtBKWtBGEpxUlpZMWtBKTmOcwDn9xDX1gDn9wjf50p5WjmOc2sw"
      Columns(2).ValueItems(7).DisplayValue(36)=   "AGswAGswAGswAM6+tf///////////////wD///////////////////////////9aWTGUYTlrQSkQ"
      Columns(2).ValueItems(7).DisplayValue(37)=   "194A9/8A7/8A5/cA9/8Q19ZaWTFrQSkYz8ZSYUJ7OAhjMBBChmtrQSmUYTlSYUJChmtaWTF7OAhr"
      Columns(2).ValueItems(7).DisplayValue(38)=   "MAB7OAhrMACUYTn39+////////////8A////////////////////////////tZZzhGFKazAAENfW"
      Columns(2).ValueItems(7).DisplayValue(39)=   "APf/AOf3APf/AOf3CN/nAOf3APf/a0EpWlkxSnFSOY5za0EpazAAUiAAazAAMZ6UazAAUiAAazAA"
      Columns(2).ValueItems(7).DisplayValue(40)=   "a0EpazAAYzAQ9+/n////////////AP///////////////////////////9bPxpyGc2swAEpxUgD3"
      Columns(2).ValueItems(7).DisplayValue(41)=   "/wD3/wD3/wD3/wD3/wD3/wD3/1pZMWswACmupSmunGswAGswAHs4CEp5WkKGa2tBKWswAGswAHs4"
      Columns(2).ValueItems(7).DisplayValue(42)=   "CGswAGswANbPxv///////////wD////////////////////////////37+eUcVqUYTlSYUIA9/8A"
      Columns(2).ValueItems(7).DisplayValue(43)=   "5/cA9/8A9/8A9/8A5/cA9/9rQSlaWTFKeVpaWTFKeVprMABKeVpChmtrQSlrMABrMABrMABrMABr"
      Columns(2).ValueItems(7).DisplayValue(44)=   "MABrQSn37+f///////////8A////////////////////////////////tZZzSnFSazAAKa6cAPf/"
      Columns(2).ValueItems(7).DisplayValue(45)=   "AOf3WlkxSnlaAPf/AOf3OY5za0EpezgISiAAWlkxMZ6UQoZrUiAAazAASiAAazAAazAAa0EYazAA"
      Columns(2).ValueItems(7).DisplayValue(46)=   "zr6t////////////////AP////////////////////////////////fv72swAIRRIWswAGtBKTGe"
      Columns(2).ValueItems(7).DisplayValue(47)=   "lGMwEDmOcwD3/wD3/zGelGswAGtBKVpZMUp5WlJhQmtBKWswAGtBKXs4CFIgAGswAJRxWtbPxv//"
      Columns(2).ValueItems(7).DisplayValue(48)=   "/////////////////wD///////////////////////////////////9aWTGUYTlrMABrMABrMACU"
      Columns(2).ValueItems(7).DisplayValue(49)=   "YTkQ19YA9/8xnpRaWTFjMBBaWTFaWTFrMABrMACUYTlrMABrMABaWTHOvq3Wz8b/////////////"
      Columns(2).ValueItems(7).DisplayValue(50)=   "//////////////8A////////////////////////////////////9+/nhGFKazAASiAAazAAUiAA"
      Columns(2).ValueItems(7).DisplayValue(51)=   "azAAa0EpazAAUiAAazAAazAAazAAUiAAazAAa0EplGE57+fe////////////////////////////"
      Columns(2).ValueItems(7).DisplayValue(52)=   "////////////AP////////////////////////////////////f375yGc1pZMXs4CGswAGswAGsw"
      Columns(2).ValueItems(7).DisplayValue(53)=   "AHs4CGswAIRJIWswAGswAGtBKZRhOa2WhP//////////////////////////////////////////"
      Columns(2).ValueItems(7).DisplayValue(54)=   "/////////wD///////////////////////////////////////+ljnuUYTlrQSlrMABSIABrMABr"
      Columns(2).ValueItems(7).DisplayValue(55)=   "QSlrMABrMABrMACljnvWz8b/////////////////////////////////////////////////////"
      Columns(2).ValueItems(7).DisplayValue(56)=   "//////8A////////////////////////////////////////9+/nrZaEtZZza0EphEkha0EpazAA"
      Columns(2).ValueItems(7).DisplayValue(57)=   "a0Epzr6t9/fv////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(7).DisplayValue(58)=   "////AP////////////////////////////////////////////fv562WhJRhOYRRIbWWc/fv7///"
      Columns(2).ValueItems(7).DisplayValue(59)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(7).DisplayValue(60)=   "/wD////////////////////////////////////////////////////39+//////////////////"
      Columns(2).ValueItems(7).DisplayValue(61)=   "//////////////////////////////////////////////////////////////////////////8A"
      Columns(2).ValueItems(7).DisplayValue(62)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(7).DisplayValue(63)=   "////////////////////////////////////////////////////////////////////////AA=="
      Columns(2).ValueItems(7).DisplayValue.vt=   9
      Columns(2).ValueItems(7)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(8)._DefaultItem=   0
      Columns(2).ValueItems(8).Value=   "009"
      Columns(2).ValueItems(8).Value.vt=   8
      Columns(2).ValueItems(8).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(2).ValueItems(8).DisplayValue(0)=   "bHQAADYOAABCTTYOAAAAAAAANgAAACgAAAAlAAAAIAAAAAEAGAAAAAAAAA4AAAAAAAAAAAAAAAAA"
      Columns(2).ValueItems(8).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(8).DisplayValue(2)=   "///////////////////37+fez8b37+f/////////////////////////////////////////////"
      Columns(2).ValueItems(8).DisplayValue(3)=   "//8A////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(8).DisplayValue(4)=   "////9+/nxqaUlGE5WjAQazAAa0EhnIZz////////////////////////////////////////////"
      Columns(2).ValueItems(8).DisplayValue(5)=   "AP////////////////////////////////////////////////////////////////////f375xp"
      Columns(2).ValueItems(8).DisplayValue(6)=   "SmtBIWswAFogAGswAGswAGswAEogALWWc////////////////////////////////////////wD/"
      Columns(2).ValueItems(8).DisplayValue(7)=   "///////////////////////////////////////////////////////37+/Ovq1rQSFrMABrMAAp"
      Columns(2).ValueItems(8).DisplayValue(8)=   "rqUYvrUA9/8xppQ5jnNrQSFrMABrQSH37+f///////////////////////////////////8A////"
      Columns(2).ValueItems(8).DisplayValue(9)=   "////////////////////////////////////////////1se1nIZzazAASiAAazAAazAAGM/GGL61"
      Columns(2).ValueItems(8).DisplayValue(10)=   "ezgIGL61WlkxENfWSo5zazAAazAA3s/G////////////////////////////////////AP//////"
      Columns(2).ValueItems(8).DisplayValue(11)=   "//////////////////////////////fv55SGc5RhOVogAGswAFogAGswAFogADmOc3tZQlpZMSmO"
      Columns(2).ValueItems(8).DisplayValue(12)=   "hBDX1imOhGswABi+tWswAFogAFpZMf///////////////////////////////////wD/////////"
      Columns(2).ValueItems(8).DisplayValue(13)=   "///////////////////ez8aUcVpKIABrMABrMAB7OAhrMABrMABrMAB7OAghrpwI3+cQ19YA7/8A"
      Columns(2).ValueItems(8).DisplayValue(14)=   "9/8Q194Q19YA9/9SOCFrMABrMAD37+f///////////////////////////////8A////////////"
      Columns(2).ValueItems(8).DisplayValue(15)=   "////////3s/Ga0EhWiAAWiAAazAAazAAazAAWjAQazAAazAAazAAe1lCOY5za0EhazAAKY6EENfW"
      Columns(2).ValueItems(8).DisplayValue(16)=   "OXFjezgIKY6EWlkxWiAAazAAlIZz////////////////////////////////AP//////////////"
      Columns(2).ValueItems(8).DisplayValue(17)=   "/+ff1mswAGswAGswAGswAGswAGswAGswAGswAGswAGswAGswAEqOc1JhQhDX3jlxY2tBGBDX1lpZ"
      Columns(2).ValueItems(8).DisplayValue(18)=   "MRi+tRjPxmswAFpZMSGunDmupf///////////////////////////////wD///////////////+U"
      Columns(2).ValueItems(8).DisplayValue(19)=   "hnNrMABrQSFrMABrMABrMABaIABrMABrQSFrMABrMABrMABSOCFrQRg5cWMQ19Y5cWMprqU5cWMQ"
      Columns(2).ValueItems(8).DisplayValue(20)=   "19ZrQSFSYUI5cWMA9/8Q19bez8b///////////////////////////8A////////////////lGE5"
      Columns(2).ValueItems(8).DisplayValue(21)=   "a0EhazAAazAAezgIa0EhazAAazAAezgIazAAazAAazAAa0EYIa6cazAAazAASo5zOXFjOY5zOXFj"
      Columns(2).ValueItems(8).DisplayValue(22)=   "GM/GAPf/Ka6lOY5zAPf/Oa6l////////////////////////////AP///////////////5SGc2tB"
      Columns(2).ValueItems(8).DisplayValue(23)=   "GFogAGswAFogAGswAGswAGswAGtBIWswAGswAGswAFogAFJhQhDX1mswAHtZQjGmlBi+tQD3/wjf"
      Columns(2).ValueItems(8).DisplayValue(24)=   "5xDX1hi+tVpZMQD3/zGmlPfv7////////////////////////wD////////////////Wx7WUcVqE"
      Columns(2).ValueItems(8).DisplayValue(25)=   "SRhrMABrMABrMAB7OAhrMACESRhrMABrMABrMABrMABSOCEA9/8I3+cA9/8A9/8A9/8Q19YYz8YQ"
      Columns(2).ValueItems(8).DisplayValue(26)=   "19YQ194I3+cA9/8A9//ez8b///////////////////////8A////////////////////lGE5e1lC"
      Columns(2).ValueItems(8).DisplayValue(27)=   "azAAazAAazAAWiAAazAAa0EhazAAazAAazAAOXFjOY5zMaaUOY5zKY6EOY5zKY6EMaaUCN/nAPf/"
      Columns(2).ValueItems(8).DisplayValue(28)=   "CN/nENfWCN/nKa6la0Eh////////////////////////AP///////////////////62ejJRhOWsw"
      Columns(2).ValueItems(8).DisplayValue(29)=   "AHs4CGswAGswAGswAHs4CGswAGswAGtBIQD3/wD3/wD3/wD3/wD3/wD3/1pZMTGmlBjPxhi+tSmu"
      Columns(2).ValueItems(8).DisplayValue(30)=   "pSGunHs4CEogAGswAN7Pxv///////////////////wD////////////////////ez8ZrQSGESRha"
      Columns(2).ValueItems(8).DisplayValue(31)=   "IABrMABrMABrMABaIABrMABaIABrMAAI3+cI3+cI3+cA9/8I3+cA9/85cWMprqUI3+cQ19YI3+cp"
      Columns(2).ValueItems(8).DisplayValue(32)=   "rqVaIABrMABaIACtnoz///////////////////8A////////////////////////lGE5c3FaazAA"
      Columns(2).ValueItems(8).DisplayValue(33)=   "SiAAWlkxOXFjOY5zIa6cKa6lc3FaAPf/APf/APf/CN/nAPf/APf/So5zOXFjAPf/APf/GM/GCN/n"
      Columns(2).ValueItems(8).DisplayValue(34)=   "APf/APf/Ka6lazAA9+/n////////////////AP///////////////////////97PxoRJGGtBIWsw"
      Columns(2).ValueItems(8).DisplayValue(35)=   "AGtBIVpZMTlxYzmOcyG2rVpZMQD3/wD3/wjf5wD3/wjf5wD3/ymOhGswAAjf5wD3/xDX1gjf5wjf"
      Columns(2).ValueItems(8).DisplayValue(36)=   "5zmOc3NxWmswAM6+tf///////////////wD///////////////////////////9aWTGUYTlrMABr"
      Columns(2).ValueItems(8).DisplayValue(37)=   "QRhSOCFaWTE5cWM5jnM5cWMA9/8I3+cA9/8A9/8A9/8A9/85jnNaIABKjnMA9/8I3+cA9/8A7/9S"
      Columns(2).ValueItems(8).DisplayValue(38)=   "YUJ7OAhrMACUcVr39+////////////8A////////////////////////////nIZzhGFKazAAUjgh"
      Columns(2).ValueItems(8).DisplayValue(39)=   "WlkxKY6EMaaUGL61a0EYAPf/APf/CN/nCN/nCN/nAPf/OXFjazAAWiAAMaaUAPf/CN/nCN/nOY5z"
      Columns(2).ValueItems(8).DisplayValue(40)=   "IbatazAAWjAQ9+/n////////////AP///////////////////////////97PxpyGc2swAGswAGtB"
      Columns(2).ValueItems(8).DisplayValue(41)=   "IVpZMTlxY0qOc2tBIRDX3gD3/wD3/wD3/wD3/wD3/1pZMWswAHs4CGswAEqOcyGunCmupQjf50qO"
      Columns(2).ValueItems(8).DisplayValue(42)=   "c1ogAGswAN7Pxv///////////wD////////////////////////////37+eUcVqcaUprQSFaWTE5"
      Columns(2).ValueItems(8).DisplayValue(43)=   "cWM5jnMxppRrMABSOCEprqUYvrUprqUYvrU5jnNaIABrMABrMABrMABaIABrMAAA9/8A9/8I3+dr"
      Columns(2).ValueItems(8).DisplayValue(44)=   "MABrQSH37+f///////////8A////////////////////////////////tZZzc3FaazAAa0EhSo5z"
      Columns(2).ValueItems(8).DisplayValue(45)=   "MaaUOY5zENfWWlkxIa6cCN/nENfWGM/GSiAAazAAazAAezgIazAAazAAazAAAO//CN/nCN/nWiAA"
      Columns(2).ValueItems(8).DisplayValue(46)=   "zr6t////////////////AP////////////////////////////////fv72swAIxRIWswAGtBIVpZ"
      Columns(2).ValueItems(8).DisplayValue(47)=   "MTlxYzmOc2tBISmupQjf5wjf5wjf52swAGswAGswAGtBIWswAGtBIXs4CFogAFpZMZRhOeff1v//"
      Columns(2).ValueItems(8).DisplayValue(48)=   "/////////////////wD///////////////////////////////////9aWTFaWTFrMACESRhrMABr"
      Columns(2).ValueItems(8).DisplayValue(49)=   "MABrMABrMABSOCFaWTFSYUJaWTFrMABrMABrMACESRhKIABrMABaWTHOvq3ez8b/////////////"
      Columns(2).ValueItems(8).DisplayValue(50)=   "//////////////8A////////////////////////////////////9+/nhGFKazAAWiAAazAAWiAA"
      Columns(2).ValueItems(8).DisplayValue(51)=   "azAAazAAazAAWiAAazAAWiAAazAAWiAAazAAa0EhlGE57+fe////////////////////////////"
      Columns(2).ValueItems(8).DisplayValue(52)=   "////////////AP////////////////////////////////////f375yGc1pZMXs4CGswAGswAGsw"
      Columns(2).ValueItems(8).DisplayValue(53)=   "AHs4CGswAIRJGGswAGswAGtBIZRhOa2ejP//////////////////////////////////////////"
      Columns(2).ValueItems(8).DisplayValue(54)=   "/////////wD///////////////////////////////////////+UhnOUYTlrQSFrMABaIABrMABr"
      Columns(2).ValueItems(8).DisplayValue(55)=   "QSFrMABrMABrMACUhnPez8b/////////////////////////////////////////////////////"
      Columns(2).ValueItems(8).DisplayValue(56)=   "//////8A////////////////////////////////////////9+/nrZ6MtZZza0EhhEkYa0EhazAA"
      Columns(2).ValueItems(8).DisplayValue(57)=   "a0Ehzr6t9/fv////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(8).DisplayValue(58)=   "////AP////////////////////////////////////////////fv58amlJRhOYxRIbWWc/fv7///"
      Columns(2).ValueItems(8).DisplayValue(59)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(8).DisplayValue(60)=   "/wD////////////////////////////////////////////////////39+//////////////////"
      Columns(2).ValueItems(8).DisplayValue(61)=   "//////////////////////////////////////////////////////////////////////////8A"
      Columns(2).ValueItems(8).DisplayValue(62)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(8).DisplayValue(63)=   "////////////////////////////////////////////////////////////////////////AA=="
      Columns(2).ValueItems(8).DisplayValue.vt=   9
      Columns(2).ValueItems(8)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(9)._DefaultItem=   0
      Columns(2).ValueItems(9).Value=   "010"
      Columns(2).ValueItems(9).Value.vt=   8
      Columns(2).ValueItems(9).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(2).ValueItems(9).DisplayValue(0)=   "bHQAADYOAABCTTYOAAAAAAAANgAAACgAAAAlAAAAIAAAAAEAGAAAAAAAAA4AAAAAAAAAAAAAAAAA"
      Columns(2).ValueItems(9).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(9).DisplayValue(2)=   "///////////////////37+fez8b37+f/////////////////////////////////////////////"
      Columns(2).ValueItems(9).DisplayValue(3)=   "//8A////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(9).DisplayValue(4)=   "////9+/nxqaUlGE5WjAQazAAa0EhnIZz////////////////////////////////////////////"
      Columns(2).ValueItems(9).DisplayValue(5)=   "AP////////////////////////////////////////////////////////////////////f375xp"
      Columns(2).ValueItems(9).DisplayValue(6)=   "SmtBIWswAFogAGswAGswAGswAEogALWWc////////////////////////////////////////wD/"
      Columns(2).ValueItems(9).DisplayValue(7)=   "///////////////////////////////////////////////////////37+/Ovq1rQSFrMABrMAAp"
      Columns(2).ValueItems(9).DisplayValue(8)=   "rqUYvrUA9/8xppQ5jnNrQSFrMABrQSH37+f///////////////////////////////////8A////"
      Columns(2).ValueItems(9).DisplayValue(9)=   "////////////////////////////////////////////1se1nIZzazAASiAAazAAazAAGM/GGL61"
      Columns(2).ValueItems(9).DisplayValue(10)=   "ezgIGL61WlkxENfWSo5zazAAazAA3s/G////////////////////////////////////AP//////"
      Columns(2).ValueItems(9).DisplayValue(11)=   "//////////////////////////////fv55SGc5RhOVogAGswAFogAGswAFogADmOc3tZQlpZMSmO"
      Columns(2).ValueItems(9).DisplayValue(12)=   "hBDX1imOhGswABi+tWswAFogAFpZMf///////////////////////////////////wD/////////"
      Columns(2).ValueItems(9).DisplayValue(13)=   "///////////////////ez8aUcVpKIABrMABrMAB7OAhrMABrMABrMAB7OAghrpwI3+cQ19YA7/8A"
      Columns(2).ValueItems(9).DisplayValue(14)=   "9/8Q194Q19YA9/9SOCFrMABrMAD37+f///////////////////////////////8A////////////"
      Columns(2).ValueItems(9).DisplayValue(15)=   "////////3s/Ga0EhWiAAWiAAazAAazAAazAAWjAQazAAazAAazAAe1lCOY5za0EhazAAKY6EENfW"
      Columns(2).ValueItems(9).DisplayValue(16)=   "OXFjezgIKY6EWlkxWiAAazAAlIZz////////////////////////////////AP//////////////"
      Columns(2).ValueItems(9).DisplayValue(17)=   "/+ff1mswAGswAGswAGswAGswAGswAGswAGswAGswAGswAGswAEqOc1JhQhDX3jlxY2tBGBDX1lpZ"
      Columns(2).ValueItems(9).DisplayValue(18)=   "MRi+tRjPxmswAFpZMSGunDmupf///////////////////////////////wD///////////////+U"
      Columns(2).ValueItems(9).DisplayValue(19)=   "hnNrMABrQSFrMABrMABrMABaIABrMABrQSFrMABrMABrMABSOCFrQRg5cWMQ19Y5cWMprqU5cWMQ"
      Columns(2).ValueItems(9).DisplayValue(20)=   "19ZrQSFSYUI5cWMA9/8Q19bez8b///////////////////////////8A////////////////lGE5"
      Columns(2).ValueItems(9).DisplayValue(21)=   "a0EhazAAazAAezgIa0EhazAAazAAezgIazAAazAAazAAa0EYIa6cazAAazAASo5zOXFjOY5zOXFj"
      Columns(2).ValueItems(9).DisplayValue(22)=   "GM/GAPf/Ka6lOY5zAPf/Oa6l////////////////////////////AP///////////////5SGc2tB"
      Columns(2).ValueItems(9).DisplayValue(23)=   "GFogAGswAFogAGswAGswAGswAGtBIWswAGswAGswAFogAFJhQhDX1mswAHtZQjGmlBi+tQD3/wjf"
      Columns(2).ValueItems(9).DisplayValue(24)=   "5xDX1hi+tVpZMQD3/zGmlPfv7////////////////////////wD////////////////Wx7WUcVqE"
      Columns(2).ValueItems(9).DisplayValue(25)=   "SRhrMABrMABrMAB7OAhrMACESRhrMABrMABrMABrMABSOCEA9/8I3+cA9/8A9/8A9/8Q19YYz8YQ"
      Columns(2).ValueItems(9).DisplayValue(26)=   "19YQ194I3+cA9/8A9//ez8b///////////////////////8A////////////////////lGE5e1lC"
      Columns(2).ValueItems(9).DisplayValue(27)=   "azAAazAAazAAWiAAazAAa0EhazAAazAAazAAOXFjOY5zMaaUOY5zKY6EOY5zKY6EMaaUCN/nAPf/"
      Columns(2).ValueItems(9).DisplayValue(28)=   "CN/nENfWCN/nKa6la0Eh////////////////////////AP///////////////////62ejJRhOWsw"
      Columns(2).ValueItems(9).DisplayValue(29)=   "AHs4CGswAGswAGswAHs4CGswAGswAGtBIQD3/wD3/wD3/wD3/wD3/wD3/1pZMTGmlBjPxhi+tSmu"
      Columns(2).ValueItems(9).DisplayValue(30)=   "pSGunHs4CEogAGswAN7Pxv///////////////////wD////////////////////ez8ZrQSGESRha"
      Columns(2).ValueItems(9).DisplayValue(31)=   "IABrMABrMABrMABaIABrMABaIABrMAAI3+cI3+cI3+cA9/8I3+cA9/85cWMprqUI3+cQ19YI3+cp"
      Columns(2).ValueItems(9).DisplayValue(32)=   "rqVaIABrMABaIACtnoz///////////////////8A////////////////////////lGE5c3FaazAA"
      Columns(2).ValueItems(9).DisplayValue(33)=   "SiAAWlkxOXFjOY5zIa6cKa6lc3FaAPf/APf/APf/CN/nAPf/APf/So5zOXFjAPf/APf/GM/GCN/n"
      Columns(2).ValueItems(9).DisplayValue(34)=   "APf/APf/Ka6lazAA9+/n////////////////AP///////////////////////97PxoRJGGtBIWsw"
      Columns(2).ValueItems(9).DisplayValue(35)=   "AGtBIVpZMTlxYzmOcyG2rVpZMQD3/wD3/wjf5wD3/wjf5wD3/ymOhGswAAjf5wD3/xDX1gjf5wjf"
      Columns(2).ValueItems(9).DisplayValue(36)=   "5zmOc3NxWmswAM6+tf///////////////wD///////////////////////////9aWTGUYTlrMABr"
      Columns(2).ValueItems(9).DisplayValue(37)=   "QRhSOCFaWTE5cWM5jnM5cWMA9/8I3+cA9/8A9/8A9/8A9/85jnNaIABKjnMA9/8I3+cA9/8A7/9S"
      Columns(2).ValueItems(9).DisplayValue(38)=   "YUJ7OAhrMACUcVr39+////////////8A////////////////////////////nIZzhGFKazAAUjgh"
      Columns(2).ValueItems(9).DisplayValue(39)=   "WlkxKY6EMaaUGL61a0EYAPf/APf/CN/nCN/nCN/nAPf/OXFjazAAWiAAMaaUAPf/CN/nCN/nOY5z"
      Columns(2).ValueItems(9).DisplayValue(40)=   "IbatazAAWjAQ9+/n////////////AP///////////////////////////97PxpyGc2swAGswAGtB"
      Columns(2).ValueItems(9).DisplayValue(41)=   "IVpZMTlxY0qOc2tBIRDX3gD3/wD3/wD3/wD3/wD3/1pZMWswAHs4CGswAEqOcyGunCmupQjf50qO"
      Columns(2).ValueItems(9).DisplayValue(42)=   "c1ogAGswAN7Pxv///////////wD////////////////////////////37+eUcVqcaUprQSFaWTE5"
      Columns(2).ValueItems(9).DisplayValue(43)=   "cWM5jnMxppRrMABSOCEprqUYvrUprqUYvrU5jnNaIABrMABrMABrMABaIABrMAAA9/8A9/8I3+dr"
      Columns(2).ValueItems(9).DisplayValue(44)=   "MABrQSH37+f///////////8A////////////////////////////////tZZzc3FaazAAa0EhSo5z"
      Columns(2).ValueItems(9).DisplayValue(45)=   "MaaUOY5zENfWWlkxIa6cCN/nENfWGM/GSiAAazAAazAAezgIazAAazAAazAAAO//CN/nCN/nWiAA"
      Columns(2).ValueItems(9).DisplayValue(46)=   "zr6t////////////////AP////////////////////////////////fv72swAIxRIWswAGtBIVpZ"
      Columns(2).ValueItems(9).DisplayValue(47)=   "MTlxYzmOc2tBISmupQjf5wjf5wjf52swAGswAGswAGtBIWswAGtBIXs4CFogAFpZMZRhOeff1v//"
      Columns(2).ValueItems(9).DisplayValue(48)=   "/////////////////wD///////////////////////////////////9aWTFaWTFrMACESRhrMABr"
      Columns(2).ValueItems(9).DisplayValue(49)=   "MABrMABrMABSOCFaWTFSYUJaWTFrMABrMABrMACESRhKIABrMABaWTHOvq3ez8b/////////////"
      Columns(2).ValueItems(9).DisplayValue(50)=   "//////////////8A////////////////////////////////////9+/nhGFKazAAWiAAazAAWiAA"
      Columns(2).ValueItems(9).DisplayValue(51)=   "azAAazAAazAAWiAAazAAWiAAazAAWiAAazAAa0EhlGE57+fe////////////////////////////"
      Columns(2).ValueItems(9).DisplayValue(52)=   "////////////AP////////////////////////////////////f375yGc1pZMXs4CGswAGswAGsw"
      Columns(2).ValueItems(9).DisplayValue(53)=   "AHs4CGswAIRJGGswAGswAGtBIZRhOa2ejP//////////////////////////////////////////"
      Columns(2).ValueItems(9).DisplayValue(54)=   "/////////wD///////////////////////////////////////+UhnOUYTlrQSFrMABaIABrMABr"
      Columns(2).ValueItems(9).DisplayValue(55)=   "QSFrMABrMABrMACUhnPez8b/////////////////////////////////////////////////////"
      Columns(2).ValueItems(9).DisplayValue(56)=   "//////8A////////////////////////////////////////9+/nrZ6MtZZza0EhhEkYa0EhazAA"
      Columns(2).ValueItems(9).DisplayValue(57)=   "a0Ehzr6t9/fv////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(9).DisplayValue(58)=   "////AP////////////////////////////////////////////fv58amlJRhOYxRIbWWc/fv7///"
      Columns(2).ValueItems(9).DisplayValue(59)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(9).DisplayValue(60)=   "/wD////////////////////////////////////////////////////39+//////////////////"
      Columns(2).ValueItems(9).DisplayValue(61)=   "//////////////////////////////////////////////////////////////////////////8A"
      Columns(2).ValueItems(9).DisplayValue(62)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(9).DisplayValue(63)=   "////////////////////////////////////////////////////////////////////////AA=="
      Columns(2).ValueItems(9).DisplayValue.vt=   9
      Columns(2).ValueItems(9)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(10)._DefaultItem=   0
      Columns(2).ValueItems(10).Value=   "011"
      Columns(2).ValueItems(10).Value.vt=   8
      Columns(2).ValueItems(10).DisplayValue=   "011"
      Columns(2).ValueItems(10).DisplayValue.vt=   8
      Columns(2).ValueItems(10)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(11)._DefaultItem=   0
      Columns(2).ValueItems(11).Value=   "012"
      Columns(2).ValueItems(11).Value.vt=   8
      Columns(2).ValueItems(11).DisplayValue=   "012"
      Columns(2).ValueItems(11).DisplayValue.vt=   8
      Columns(2).ValueItems(11)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(12)._DefaultItem=   0
      Columns(2).ValueItems(12).Value=   "013"
      Columns(2).ValueItems(12).Value.vt=   8
      Columns(2).ValueItems(12).DisplayValue=   "013"
      Columns(2).ValueItems(12).DisplayValue.vt=   8
      Columns(2).ValueItems(12)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(13)._DefaultItem=   0
      Columns(2).ValueItems(13).Value=   "014"
      Columns(2).ValueItems(13).Value.vt=   8
      Columns(2).ValueItems(13).DisplayValue=   "014"
      Columns(2).ValueItems(13).DisplayValue.vt=   8
      Columns(2).ValueItems(13)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(14)._DefaultItem=   0
      Columns(2).ValueItems(14).Value=   "015"
      Columns(2).ValueItems(14).Value.vt=   8
      Columns(2).ValueItems(14).DisplayValue=   "015"
      Columns(2).ValueItems(14).DisplayValue.vt=   8
      Columns(2).ValueItems(14)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(15)._DefaultItem=   0
      Columns(2).ValueItems(15).Value=   "016"
      Columns(2).ValueItems(15).Value.vt=   8
      Columns(2).ValueItems(15).DisplayValue=   "016"
      Columns(2).ValueItems(15).DisplayValue.vt=   8
      Columns(2).ValueItems(15)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(16)._DefaultItem=   0
      Columns(2).ValueItems(16).Value=   "017"
      Columns(2).ValueItems(16).Value.vt=   8
      Columns(2).ValueItems(16).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(2).ValueItems(16).DisplayValue(0)=   "bHQAAP4KAABCTf4KAAAAAAAANgAAACgAAAAeAAAAHgAAAAEAGAAAAAAAyAoAAAAAAAAAAAAAAAAA"
      Columns(2).ValueItems(16).DisplayValue(1)=   "AAAAAAD////////////////////////////////////////////////////////////////////W"
      Columns(2).ValueItems(16).DisplayValue(2)=   "x8bOvrVrOCmchnvWx8b///////////////////////////////8AAP//////////////////////"
      Columns(2).ValueItems(16).DisplayValue(3)=   "/////////////////////////////////+ff3r2mnGtJQjkQADkQAFogEDkQADkQAJyGe///////"
      Columns(2).ValueItems(16).DisplayValue(4)=   "/////////////////////wAA////////////////////////////////////////////////7+/n"
      Columns(2).ValueItems(16).DisplayValue(5)=   "azgpazgpORAAQhAIORAAWiAQORAAWiAQORAAORAA////////////////////////////AAD/////"
      Columns(2).ValueItems(16).DisplayValue(6)=   "///////////////////////////////v7+etlpRaIBA5EABaIBBaIBA5EABCGBBaIBBCGBBaIBBa"
      Columns(2).ValueItems(16).DisplayValue(7)=   "IBBaIBA5EACchnv///////////////////////8AAP///////////////////////////62WlJRp"
      Columns(2).ValueItems(16).DisplayValue(8)=   "YzkQADkQADkQAFogEDkQAFogEDkQAFogEDkQAFogEDkQAFogEDkQAFogEEI4Kf//////////////"
      Columns(2).ValueItems(16).DisplayValue(9)=   "/////////wAA////////////////59/enIZ7a0lCORAAORAAUhgIORAAUhgIUhgIORAAORAAUhgI"
      Columns(2).ValueItems(16).DisplayValue(10)=   "ORAAORAAORAAUhgIQhgQWiAQORAAWiAQxra1////////////////////AAD////////OvrVrOCla"
      Columns(2).ValueItems(16).DisplayValue(11)=   "IBA5EABSGAhCOCkxaWMYtr0Yx84Qz9YYx84Qz9YYx84Qz9YYx84I5/cYx84xaWNaIBA5EABaIBA5"
      Columns(2).ValueItems(16).DisplayValue(12)=   "EACUaWP///////////////////8AAP///2tJQjkQAFogEDkQAFIYCGtJQgjv/wjn9ymWlBi2vQjv"
      Columns(2).ValueItems(16).DisplayValue(13)=   "/wjn9xjHzgjn9wjv/xi2vTF5exi2vQjv/wjv/zFpY1ogEDkQAEI4Keff3v///////////////wAA"
      Columns(2).ValueItems(16).DisplayValue(14)=   "zr61ORAAWiAQQhAIUhgIOVFSCO//EM/Wa0lCORAAGMfOEM/We2FaQhgQa0lCCOf3Kba9ORAAWiAQ"
      Columns(2).ValueItems(16).DisplayValue(15)=   "QjgpKba9CO//MXl7ORAAORAArZaU////////////////AACchntSGAhCGBA5EABCGBAYx84I7/9a"
      Columns(2).ValueItems(16).DisplayValue(16)=   "IBBSGAgxeXsI7/85UUo5EABSGAhSGAgxeXsI7/9rOCk5EABaIBBaIBAplpQI7/9rSUI5EABrOCn/"
      Columns(2).ValueItems(16).DisplayValue(17)=   "//////////////8AAJyGe0I4KVogEDkQADlRSgjv/ymWlDkQAFogEBi2vQjv/0I4KVogEDkQAFog"
      Columns(2).ValueItems(16).DisplayValue(18)=   "EDlRUgjv/zFpY1ogEDkQAFogEEI4KRjHzhDP1logEDkQAOff3v///////////wAAxra1a0lCWiAQ"
      Columns(2).ValueItems(16).DisplayValue(19)=   "WiAQKYaECO//QjgpORAAORAAa0lCCO//Kba9OVFSWiAQORAAMWljCOf3a0lCWiAQUhgIQhgQUhgI"
      Columns(2).ValueItems(16).DisplayValue(20)=   "KYaECO//QjgpORAAnIZ7////////////AADv7+d7YVprOCk5EAAplpQI5/daIBA5EABaIBA5EAB7"
      Columns(2).ValueItems(16).DisplayValue(21)=   "YVoI7/8I7/8hpqUplpQQz9YI7/8plpRaIBA5EABaIBA5EABrSUII7/97YVo5EABaIBD/////////"
      Columns(2).ValueItems(16).DisplayValue(22)=   "//8AAPfv75RpY2s4KVogECm2vRDP1kIYEFogEEIYEFIYCDkQAFogECmGhBDP1hDP1gjn9wjn9wjv"
      Columns(2).ValueItems(16).DisplayValue(23)=   "/xi2vUIQCFogEDkQAEIYEAjv/ymGhDkQADkQAM6+tf///////wAA////rZaUlGljQhAIGMfOCO//"
      Columns(2).ValueItems(16).DisplayValue(24)=   "GMfOGLa9GMfOEM/WGMfOGLa9Kba9IaalKba9COf3CO//CO//CO//CO//Kba9GLa9GMfOCO//Kba9"
      Columns(2).ValueItems(16).DisplayValue(25)=   "ORAAWiAQazgp////////AAD////v7+drSUJrOCkhpqUptr0Ytr0ptr0Ytr0ptr0Ytr0ptr0Ytr0p"
      Columns(2).ValueItems(16).DisplayValue(26)=   "tr0Ytr0plpQYtr0I7/8I5/cplpQI7/8I7/8Ytr0plpQphoRaIBBCGBBaIBDn397///8AAP//////"
      Columns(2).ValueItems(16).DisplayValue(27)=   "/5yGe3thWlogEDkQAFogEDkQAFogEDkQAFogEDkQAFogEDkQAFogEDkQAFogECmOjAjv/zFpY2s4"
      Columns(2).ValueItems(16).DisplayValue(28)=   "KQjn9wjv/2s4KVogEDkQAFogEDkQAL2mnP///wAA////////59/eazgpWiAQORAAWiAQUhgIQhgQ"
      Columns(2).ValueItems(16).DisplayValue(29)=   "WiAQORAAWiAQWiAQORAAORAAWiAQQhgQORAAIaalCO//OVFSWiAQGLa9CO//KYaEWiAQORAAWiAQ"
      Columns(2).ValueItems(16).DisplayValue(30)=   "a0lC////AAD///////////97YVqEWUo5EABaIBA5EABaIBBCEAhaIBA5EABaIBA5EABaIBA5EABa"
      Columns(2).ValueItems(16).DisplayValue(31)=   "IBA5EABaIBAhpqUI7/9COClaIBApjowI7/8Qz9ZrSUI5EABaIBDWx8YAAP///////////86+tWtJ"
      Columns(2).ValueItems(16).DisplayValue(32)=   "QlogEEIYEEIQCEIYEFogEEIYEFogEEIYEEIQCEIYEFogEEIYEFogEEIYEFIYCAjn9xDP1jlRUlIY"
      Columns(2).ValueItems(16).DisplayValue(33)=   "CDlRUgjn9wjn91ogEDkQAM6+tQAA////////////////hFlKazgpWiAQORAAWiAQQhgQWiAQORAA"
      Columns(2).ValueItems(16).DisplayValue(34)=   "WiAQORAAWiAQQhAIWiAQORAAWiAQORAAWiAQCOf3GMfOORAAazgpKY6MCO//UhgIORAAtaalAAD/"
      Columns(2).ValueItems(16).DisplayValue(35)=   "//////////////+chntrOClCGBBaIBBCGBA5EABCGBBaIBBCGBBaIBBCGBBaIBBCGBBaIBA5EABa"
      Columns(2).ValueItems(16).DisplayValue(36)=   "IBBaIBA5UUoI5/cYx84phoQQz9YI7/9SGAhCGBDv7+cAAP///////////////86+tWs4KVogEDkQ"
      Columns(2).ValueItems(16).DisplayValue(37)=   "AFogEDkQAFogEDkQAFogEDkQAFogEDkQAFogEEIQCFogEDkQAFogEDkQAGtJQhDP1gjv/wjv/3th"
      Columns(2).ValueItems(16).DisplayValue(38)=   "WjkQAM6+tf///wAA////////////////9+/vhFlKa0lCORAAQhgQWiAQORAAQhAIWiAQWiAQQhgQ"
      Columns(2).ValueItems(16).DisplayValue(39)=   "WiAQORAAQhAIWiAQWiAQQhgQWiAQWiAQUhgIWiAQWiAQlGlj////////////AAD/////////////"
      Columns(2).ValueItems(16).DisplayValue(40)=   "//////+tlpSUaWM5EABaIBBCGBBaIBBCEAhaIBA5EABaIBA5EABaIBA5EABaIBA5EABaIBA5EABa"
      Columns(2).ValueItems(16).DisplayValue(41)=   "IBCUaWPOvrX///////////////////8AAP///////////////////9bHxnthWlogEEIYEFogEDkQ"
      Columns(2).ValueItems(16).DisplayValue(42)=   "AFogEEIYEDkQAEIYEFogEEIYEFogEFogEEIQCGs4Ka2WlNbHxv//////////////////////////"
      Columns(2).ValueItems(16).DisplayValue(43)=   "/wAA////////////////////////hFlKe2FaWiAQQhAIWiAQORAAWiAQORAAWiAQORAAWiAQQjgp"
      Columns(2).ValueItems(16).DisplayValue(44)=   "nIZ71sfG////////////////////////////////////////AAD////////////////////////n"
      Columns(2).ValueItems(16).DisplayValue(45)=   "395rSUJrOClaIBBCGBBaIBBCGBBaIBBCGBBrSUKtlpT/////////////////////////////////"
      Columns(2).ValueItems(16).DisplayValue(46)=   "//////////////////8AAP///////////////////////////62WlK2OhGs4KVogEDkQAIRZSpRp"
      Columns(2).ValueItems(16).DisplayValue(47)=   "Y9bHxv///////////////////////////////////////////////////////////wAA////////"
      Columns(2).ValueItems(16).DisplayValue(48)=   "////////////////////////9+/vvaacnIZ7zr619+/v////////////////////////////////"
      Columns(2).ValueItems(16).DisplayValue(49)=   "////////////////////////////////////AAA="
      Columns(2).ValueItems(16).DisplayValue.vt=   9
      Columns(2).ValueItems(16)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(17)._DefaultItem=   0
      Columns(2).ValueItems(17).Value=   "019"
      Columns(2).ValueItems(17).Value.vt=   8
      Columns(2).ValueItems(17).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(2).ValueItems(17).DisplayValue(0)=   "bHQAADYOAABCTTYOAAAAAAAANgAAACgAAAAlAAAAIAAAAAEAGAAAAAAAAA4AAAAAAAAAAAAAAAAA"
      Columns(2).ValueItems(17).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(17).DisplayValue(2)=   "///////////////////v7+/e39bv7+f/////////////////////////////////////////////"
      Columns(2).ValueItems(17).DisplayValue(3)=   "//8A////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(17).DisplayValue(4)=   "////////xr6tjGlSYzAYWiAAWigQpY57////////////////////////////////////////////"
      Columns(2).ValueItems(17).DisplayValue(5)=   "AP////////////////////////////////////////////////////////////////////f375R5"
      Columns(2).ValueItems(17).DisplayValue(6)=   "a2MwKVogAGMgAGsoAHMoAGsoAFIYAKWOhP///////////////////////////////////////wD/"
      Columns(2).ValueItems(17).DisplayValue(7)=   "///////////////////////////////////////////////////////39/etx8Z7SSlSCABaGABr"
      Columns(2).ValueItems(17).DisplayValue(8)=   "KABzMABzMABzMABzMABzMABrKABjMBj39/f///////////////////////////////////8A////"
      Columns(2).ValueItems(17).DisplayValue(9)=   "////////////////////////////////////////////zse9c8/OCOfvAP//MXFSa0EQczAAczAA"
      Columns(2).ValueItems(17).DisplayValue(10)=   "czAAczAAczAAczAAczAAczAAShAAzr61////////////////////////////////////AP//////"
      Columns(2).ValueItems(17).DisplayValue(11)=   "/////////////////////////////+/v57WmlHtRQlogAEJJMTGWhBDXzgD//yG2rUKGa3sgAHMw"
      Columns(2).ValueItems(17).DisplayValue(12)=   "AHMwAHMwAHMwAHMwAHMwAFogAGtBKf///////////////////////////////////wD/////////"
      Columns(2).ValueItems(17).DisplayValue(13)=   "///////////////////e19aEaVJaIABaGABrIABzMAB7KAB7IABzKABjSSFCeWMA9/9ChnN7IABz"
      Columns(2).ValueItems(17).DisplayValue(14)=   "MABzMABzMABzMABzMABzMABKEAD39/f///////////////////////////////8A////////////"
      Columns(2).ValueItems(17).DisplayValue(15)=   "////////3tfGa0EhQgAAYxgAczAAczAAczAAczAAczAAczAAczAAcygAjBAASnFSGL69UmFCczAA"
      Columns(2).ValueItems(17).DisplayValue(16)=   "czAAczAAczAAczAAczAAWhgAnIZz////////////////////////////////AP//////////////"
      Columns(2).ValueItems(17).DisplayValue(17)=   "/97XzmMwEGMgAHMwAHMwAHMwAHMwAHMwAHMwAHMwAHMwAHMwAHMwAHsgAEKGa0pxUkpxWkppSmNJ"
      Columns(2).ValueItems(17).DisplayValue(18)=   "IXsoAHMwAHMwAHMwAGsoAHNJMff39////////////////////////////wD///////////////+l"
      Columns(2).ValueItems(17).DisplayValue(19)=   "hnNjGABzMABzMABzMABzMABzMABzMABzMABzMABzMABzMABzMAB7IAA5jnM5lns5lns5loRCeWNj"
      Columns(2).ValueItems(17).DisplayValue(20)=   "SSFzKABzMABzMABzMABaIADOvrX///////////////////////////8A////////////////nGlK"
      Columns(2).ValueItems(17).DisplayValue(21)=   "ezgIczAAczAAczAAczAAazgYazAIeygAeygAcygAczAAczAAeyAAOZZ7OYZrAPf/MZaEOY5zQoZr"
      Columns(2).ValueItems(17).DisplayValue(22)=   "azgYcygAczAAczAAWhgAlHlr////////////////////////////AP///////////////62Oa4xZ"
      Columns(2).ValueItems(17).DisplayValue(23)=   "MXMwAHMwAHMwAHMoABjHvVJpSkpxWlpZOWtBEHsoAHMoAIQYADGejEKGawjv7wjv70KGazmOczGW"
      Columns(2).ValueItems(17).DisplayValue(24)=   "hHMwAHMoAHMwAGsoAFooCPfv7////////////////////////wD////////////////Wvq2laUKE"
      Columns(2).ValueItems(17).DisplayValue(25)=   "MABzKAB7IABSaUoI395ChmsxnoxCeWMI1945hmtjSSFzKAApppRChmsA7+8A//8hvrVChms5jnMp"
      Columns(2).ValueItems(17).DisplayValue(26)=   "rqV7KABzKABzMABKCADGtq3///////////////////////8A////////////////////UratKb61"
      Columns(2).ValueItems(17).DisplayValue(27)=   "UllCa0Epa0EQY1EpSnFaMZ6EGMfGSnFSGMfGEM/OczAYKaacQoZrAO/3APf/AP//MZaEOY57MZ6E"
      Columns(2).ValueItems(17).DisplayValue(28)=   "Ka6leygAczAAYyAAa0Eh////////////////////////AP///////////////////86ehHuGazGW"
      Columns(2).ValueItems(17).DisplayValue(29)=   "hCmmlDGehEJ5a1pRMXMwAFpROWs4CCGupSmmnIwIACG2rUKGawDv9wD3/wD3/wD3/0pxWjGWhBjH"
      Columns(2).ValueItems(17).DisplayValue(30)=   "vTGejIQYAHMwAFIYAOff3v///////////////////wD////////////////////n39acYTGMOBBz"
      Columns(2).ValueItems(17).DisplayValue(31)=   "KABjSSlKcVIpppQI194xppxzKAB7KABrOAiMAAAYvr1ChmsA9/8A9/8A9/8A//8I5+daWTE5jnsA"
      Columns(2).ValueItems(17).DisplayValue(32)=   "9/9SaUqEGABSGACtnoz///////////////////8A////////////////////////lGFCnGlKczAA"
      Columns(2).ValueItems(17).DisplayValue(33)=   "cygAeygAhBgAhBAAWlk5EM/OIa6lSnFSY1EhEM/OSnFaAPf/APf/APf/APf/AP//GMfGQnljSnFS"
      Columns(2).ValueItems(17).DisplayValue(34)=   "AP//SnFacyAAYzAY9/f3////////////////AP///////////////////////97HvYRJIXtBEHMw"
      Columns(2).ValueItems(17).DisplayValue(35)=   "AHMwAHMwAHMwAHsgAHsYAEKGa0J5Y0pxWlpROUpxUgD//wD3/wD3/wD3/wD3/wD3/zmWhCmmnDGe"
      Columns(2).ValueItems(17).DisplayValue(36)=   "hAD//2tBGFoYAMa2pf///////////////wD////////////////////////39++caUKMUSFrKABz"
      Columns(2).ValueItems(17).DisplayValue(37)=   "MABzMABzMABzMAB7KABKaUoA7/cYz8Y5jnsQ19YA//8A9/8A9/8A9/8A9/8A//8Q19ZKcVoYvr1K"
      Columns(2).ValueItems(17).DisplayValue(38)=   "cVJzOAhaGACMaWP///////////////8A////////////////////////////vaaElGE5aygAczAA"
      Columns(2).ValueItems(17).DisplayValue(39)=   "czAAczAAczAAczAAeyAAOYZzAP//AP//AP//CN/nAOfvAPf/APf/AP//GL69OYZrMZaEY0kheyAA"
      Columns(2).ValueItems(17).DisplayValue(40)=   "czAAaygAYzAQ7+fn////////////AP///////////////////////////97PvaV5Wns4CHMoAHMw"
      Columns(2).ValueItems(17).DisplayValue(41)=   "AHMwAHMwAHMwAHMwAHsoADmWhAD//xDPziG+rUKGawD//wD39ymunDmOczmWe2s4AHsgAHMwAHMw"
      Columns(2).ValueItems(17).DisplayValue(42)=   "AHMoAFooAO/n5////////////wD////////////////////////////39++teWOUYTlrKABzMABz"
      Columns(2).ValueItems(17).DisplayValue(43)=   "MABzMABzMABzMABzKAB7GAAhtqUQz84hvrUxnowI3+cxloQpppQxppRaUSlzKABzMABzMABzMABj"
      Columns(2).ValueItems(17).DisplayValue(44)=   "IAB7STn39/f///////////8A////////////////////////////////tZZzpXlSczAAczAAczAA"
      Columns(2).ValueItems(17).DisplayValue(45)=   "czAAczAAczAAczAAcygAcygAENfWGM/GCO/vMZaEGL69OZZ7QoZjcygAczAAczAAczAAczAAYyAA"
      Columns(2).ValueItems(17).DisplayValue(46)=   "zr6t////////////////AP///////////////////////////////////4RJGIxZMWsoAHMwAHMw"
      Columns(2).ValueItems(17).DisplayValue(47)=   "AHMwAHMwAHMwAHMwAHsoAGNBGBDX1kJ5Yxi+vVJhOSmelIQgAHs4CHMwAHMwAIRJIZRpSuff1v//"
      Columns(2).ValueItems(17).DisplayValue(48)=   "/////////////////wD///////////////////////////////////+caUqUYTlzMABzMABzMABz"
      Columns(2).ValueItems(17).DisplayValue(49)=   "MABzMABzMABzMABzMABzMABrOBBrOBBaUSlzKAB7QRBzMABjGAB7SRjGppTe187/////////////"
      Columns(2).ValueItems(17).DisplayValue(50)=   "//////////////8A////////////////////////////////////59fOjFEpe0EQczAAczAAczAA"
      Columns(2).ValueItems(17).DisplayValue(51)=   "czAAczAAczAAczAAczAAczAAczAAezAAezgIczAAnHFK59/W////////////////////////////"
      Columns(2).ValueItems(17).DisplayValue(52)=   "////////////AP////////////////////////////////////f396V5WoRJIWsoAHMwAHMwAHMw"
      Columns(2).ValueItems(17).DisplayValue(53)=   "AHMwAHMwAHM4CHM4AHMwAHM4CJRhOb2mjPfv5///////////////////////////////////////"
      Columns(2).ValueItems(17).DisplayValue(54)=   "/////////wD///////////////////////////////////////+9poyleVpzMABrKABrKABzMAB7"
      Columns(2).ValueItems(17).DisplayValue(55)=   "OAhzOAhzMACESRitjnPWx73/////////////////////////////////////////////////////"
      Columns(2).ValueItems(17).DisplayValue(56)=   "//////8A////////////////////////////////////////7+/ntZZztY5zlFkxhEkYczAAaygA"
      Columns(2).ValueItems(17).DisplayValue(57)=   "jFkxxrac9/f3////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(17).DisplayValue(58)=   "////AP/////////////////////////////////////////////397WWe5xpSoRRKaV5Y/fv7///"
      Columns(2).ValueItems(17).DisplayValue(59)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(17).DisplayValue(60)=   "/wD/////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(17).DisplayValue(61)=   "//////////////////////////////////////////////////////////////////////////8A"
      Columns(2).ValueItems(17).DisplayValue(62)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(17).DisplayValue(63)=   "////////////////////////////////////////////////////////////////////////AA=="
      Columns(2).ValueItems(17).DisplayValue.vt=   9
      Columns(2).ValueItems(17)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems.Count=   18
      Columns(2).Caption=   "COD_MENU"
      Columns(2).DataField=   "COD_MODALIDAD_VENTA"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=529"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=450"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=5583"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5503"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1005"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=926"
      Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.bgcolor=&HFFFFFF&"
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
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.bgcolor=&HFFFFFF&,.fgcolor=&H0&"
      _StyleDefs(41)  =   ":id=32,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(42)  =   ":id=32,.fontname=MS Sans Serif"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(50)  =   "Named:id=33:Normal"
      _StyleDefs(51)  =   ":id=33,.parent=0"
      _StyleDefs(52)  =   "Named:id=34:Heading"
      _StyleDefs(53)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(54)  =   ":id=34,.wraptext=-1"
      _StyleDefs(55)  =   "Named:id=35:Footing"
      _StyleDefs(56)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(57)  =   "Named:id=36:Selected"
      _StyleDefs(58)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(59)  =   "Named:id=37:Caption"
      _StyleDefs(60)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(61)  =   "Named:id=38:HighlightRow"
      _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&HFEEBDE&,.fgcolor=&H80000012&"
      _StyleDefs(63)  =   "Named:id=39:EvenRow"
      _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(65)  =   "Named:id=40:OddRow"
      _StyleDefs(66)  =   ":id=40,.parent=33"
      _StyleDefs(67)  =   "Named:id=41:RecordSelector"
      _StyleDefs(68)  =   ":id=41,.parent=34"
      _StyleDefs(69)  =   "Named:id=42:FilterBar"
      _StyleDefs(70)  =   ":id=42,.parent=33"
   End
   Begin ORADCLibCtl.ORADC oradcModalidad 
      Height          =   255
      Left            =   2760
      Top             =   4800
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   450
      _StockProps     =   207
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DatabaseName    =   ""
      Connect         =   ""
      RecordSource    =   ""
   End
   Begin MSComctlLib.ImageList ilsImagenes 
      Left            =   2640
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   65
      ImageHeight     =   56
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Modalidad.frx":0029
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Modalidad.frx":0795
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Modalidad.frx":1054
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Modalidad.frx":183B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Modalidad.frx":209E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Modalidad.frx":2943
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Modalidad.frx":321D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Modalidad.frx":3AA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Modalidad.frx":41FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_VTA_Modalidad.frx":4AC8
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_VTA_Modalidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objModalidad As New clsModalidad
Dim objDocumento As New clsDocumento
Dim odynR1 As oraDynaset
Dim odynClon As oraDynaset
Private lblnVentana As Boolean
Public pstrCodModa As String


Private Sub Form_Load()
    lblnVentana = False
    Dim objPermisos As New clsAutorizacion
    'Set odynR1 = gclsOracle.FN_ORADC("BTLPROD.PKG_MODALIDAD.FN_LISTA", oradcModalidad, 0)
    Set oradcModalidad.Recordset = objPermisos.ListaPermisos(objUsuario.Aplicacion, objUsuario.Codigo, "076")
    'grdModalidad.DataSource = objPermisos.ListaPermisos(objUsuario.Aplicacion, objUsuario.Codigo, "076")
    Call SeteaGrilla
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Control
    'psub_KeyDownAplicacion KeyCode, Shift
    Select Case KeyCode
        Case vbKeyEscape
            'Unload Me
    End Select
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, App.ProductName
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'   ' If objVenta.CodigoTipoVenta < 1 Then Cancel = 1
'End Sub

Private Sub grdModalidad_DblClick()
If grdModalidad.ApproxCount <= 0 Then Exit Sub
On Error GoTo Control
Dim strMensaje As String
    If objVenta.ErrorFracciones(grdModalidad.Columns(0).Value, strMensaje) = True And grdModalidad.Columns(0).Value <> objVenta.CodModalidadVenta Then
        MsgBox "Los siguientes productos no fracciona para la modalidad de venta: " & grdModalidad.Columns(1).Value & Chr(13) & strMensaje, vbCritical, App.ProductName
        Exit Sub
    End If
    pstrCodModa = Trim(grdModalidad.Columns(0).Value)
    objVenta.CodModalidadVenta = pstrCodModa
    codModalidadVentaBK = pstrCodModa
    
    Select Case grdModalidad.Columns(0).Value
        Case "001"
            objVenta.NumMaximoUnidades = 0
            Unload Me
            objVenta.LimpiaServicio
            objVenta.LimpiaConvenio
            LimpiarSiSalgodeGuia
            frmPedido.grdPedido.Rebind
            If objVenta.ptmModalidad = Venta_Convenio Then objVenta.CodigoConvenio = ""
            objVenta.ptmModalidad = Venta_Regular
            ptmTipoPrecio = Regular
            objVenta.CodigoTipoVenta = Venta_Regular
            objVenta.PctBeneficiario = 0
            frmPedido.lblModalidad.Caption = objVenta.NombreTipoVenta
            'frmPedido.Label6.Visible = True
            'frmPedido.lblTotal.Visible = True
            frmPedido.Label4.Visible = False
            frmPedido.lblPctCopago.Visible = False
            frmPedido.Label8.Visible = False
            frmPedido.lblcopago.Visible = False
            frmPedido.lblSiguiente.Caption = objUsuario.TipDocDefault & " - " & objDocumento.ListaNumeroDisponible(objUsuario.CodigoEmpresa, objUsuario.NombrePC, objUsuario.TipDocDefault)
            frmPedido.Cal_Montos
            mdiPrincipal.cmdGrabaVenta.Enabled = True
            frm_VTA_RecetarioM.pstrFlgRM = ""
            Select Case objUsuario.TipoMaquina
                Case objUsuario.TipoMaquinaAdmin
                   'If ptmTipoPrecio = "3" Then Exit Sub
                    frm_VTA_Busqueda.Datos Format(ptmTipoPrecio, "000")
                Case objUsuario.TipoMaquinaCajero
                    frm_VTA_Busqueda.Datos Format(ptmTipoPrecio, "000")
                Case objUsuario.TipoMaquinaCabina
                  '** Esta parte se cambio porque no refrescaba la modalidad de la venta cuando estabas en cabina **'
                  '** Hecho 30/10/2007 Por Cristhian Rueda **'
                  '** Cambiado para cuando sea DLV el precio tega tipo "003" 22/01/2008 **'
                    If objUsuario.EsDelivery And objUsuario.flgDeliveryProv = 0 Then
                       frm_VTA_Busqueda.Datos Format("3", "000")
                    ElseIf objUsuario.EsDelivery = True And objUsuario.flgDeliveryProv = 1 Then
                       frm_VTA_Busqueda.Datos Format(ptmTipoPrecio, "000")
                    End If
                Case objUsuario.TipoMaquinaRuteo
            End Select
            If objUsuario.EsDelivery = True And mdiPrincipal.txtDLVTelefono.Visible = True Then
                frm_VTA_Busqueda.Hide
                mdiPrincipal.txtDLVTelefono.SetFocus
                mdiPrincipal.txtDLVTelefono.selection
            End If
            frmPedido.Cal_Promo
        Case "002"
           objVenta.LimpiaServicio
           LimpiarSiSalgodeGuia
            frmPedido.grdPedido.Rebind
            Unload Me
            ptmTipoPrecio = Convenio
            objVenta.CodigoTipoVenta = Venta_Convenio
            frmPedido.lblModalidad.Caption = objVenta.NombreTipoVenta
            objVenta.ptmModalidad = Venta_Convenio
            frmPedido.Label4.Visible = True
            frmPedido.lblPctCopago.Visible = True
            frmPedido.Label8.Visible = True
            frmPedido.lblcopago.Visible = True
            'frm_VTA_Busqueda.Datos Format(ptmTipoPrecio, "000") 'ECASTILLO 22.06.2020
            If objUsuario.EsDelivery And objUsuario.flgDeliveryProv = 0 Then
               frm_VTA_Busqueda.Datos Format("3", "000")
            Else
               frm_VTA_Busqueda.Datos Format(ptmTipoPrecio, "000")
            End If
            frm_VTA_RecetarioM.pstrFlgRM = ""
            mdiPrincipal.cmdGrabaVenta.Enabled = True
            frm_VTA_Convenio.Show
            frm_VTA_Convenio.SetFocus
            'limpiar variables usadas para ruteo automatico
            'objVenta.bk_codLocal = ""
            'objVenta.bk_codLocalCapacidad = ""
            'objVenta.bk_FechaCapacidad = ""
            'objVenta.bk_HoraCapacidad = ""
            'objVenta.bk_HoraCapacidad2 = ""
            'objVenta.bk_codBeneficiario = ""
            'objVenta.bk_strProforma = ""
            'objVenta.bk_ServiceType = ""
            'objVenta.bk_chkRET = ""
            'objVenta.bk_flgPactado = ""
            
        Case "003" 'Mayorista
            Unload Me
            objVenta.NumMaximoUnidades = 0
            objVenta.LimpiaServicio
            objVenta.LimpiaConvenio
            LimpiarSiSalgodeGuia
            frmPedido.grdPedido.Rebind
            If objVenta.ptmModalidad = Venta_Convenio Then objVenta.CodigoConvenio = ""
            objVenta.LimpiaServicio
            objVenta.ptmModalidad = Venta_Mayorista
            'objVenta.CodModalidadVenta = Venta_Mayorista
            ptmTipoPrecio = Mayorista
            objVenta.CodigoTipoVenta = Venta_Mayorista
            objVenta.PctBeneficiario = 0
            frmPedido.lblModalidad.Caption = objVenta.NombreTipoVenta
            'frmPedido.Label6.Visible = True
            'frmPedido.lblTotal.Visible = True
            frmPedido.Label4.Visible = False
            frmPedido.lblPctCopago.Visible = False
            frmPedido.Label8.Visible = False
            frmPedido.lblcopago.Visible = False
            frm_VTA_RecetarioM.pstrFlgRM = ""
            mdiPrincipal.cmdGrabaVenta.Enabled = True
            frm_VTA_Busqueda.Datos Format(ptmTipoPrecio, "000")
            
            frmPedido.Cal_Promo
        
        Case "004"
            Unload Me
            objVenta.NumMaximoUnidades = 0
            objVenta.LimpiaServicio
            objVenta.LimpiaConvenio
            LimpiarSiSalgodeGuia
            frmPedido.grdPedido.Rebind
            If objVenta.ptmModalidad = Venta_Convenio Then objVenta.CodigoConvenio = ""
            objVenta.ptmModalidad = Cobro_Responsabilidad
            objVenta.CodigoTipoVenta = Cobro_Responsabilidad
            objVenta.PctBeneficiario = 0
            frmPedido.lblModalidad.Caption = objVenta.NombreTipoVenta
            frmPedido.Label4.Visible = False
            frmPedido.lblPctCopago.Visible = False
            frmPedido.Label8.Visible = False
            frmPedido.lblcopago.Visible = False
            ptmTipoPrecio = Regular
            frm_VTA_Busqueda.Datos Format(ptmTipoPrecio, "000")
            frm_VTA_RecetarioM.pstrFlgRM = ""
            mdiPrincipal.cmdGrabaVenta.Enabled = False
            frm_VTA_CobroXResponsabilidad.Show
            frm_VTA_CobroXResponsabilidad.SetFocus
        Case "005"
            Unload Me
            objVenta.LimpiaServicio
            objVenta.LimpiaConvenio
            LimpiarSiSalgodeGuia
            frmPedido.grdPedido.Rebind
            objVenta.ptmModalidad = Canje_Canje
            objVenta.CodigoTipoVenta = Canje_Canje
            objVenta.PctBeneficiario = 0
            frmPedido.lblModalidad.Caption = objVenta.NombreTipoVenta
            'frmPedido.Label6.Visible = True
            'frmPedido.lblTotal.Visible = True
            frmPedido.Label4.Visible = False
            frmPedido.lblPctCopago.Visible = False
            frmPedido.Label8.Visible = False
            frmPedido.lblcopago.Visible = False
            frm_VTA_RecetarioM.pstrFlgRM = ""
            mdiPrincipal.cmdGrabaVenta.Enabled = True
            frm_VTA_Canje.Show
            frm_VTA_Canje.SetFocus
        Case "006"
            Unload Me
            'objVenta.LimpiaProyecto
            'objVenta.LimpiaServicio
            objVenta.LimpiaProductos
            objVenta.LimpiaConvenio
            LimpiarSiSalgodeGuia
            frmPedido.grdPedido.Rebind
            objVenta.ptmModalidad = Servicio
            objVenta.CodigoTipoVenta = Servicio
            objVenta.PctBeneficiario = 0
            frmPedido.lblModalidad.Caption = objVenta.NombreTipoVenta
            'frmPedido.Label6.Visible = True
            'frmPedido.lblTotal.Visible = True
            frmPedido.Label4.Visible = False
            frmPedido.lblPctCopago.Visible = False
            frmPedido.Label8.Visible = False
            frmPedido.lblcopago.Visible = False
            frmPedido.Cal_Montos
            frm_VTA_RecetarioM.pstrFlgRM = ""
            mdiPrincipal.cmdGrabaVenta.Enabled = True
            frm_VTA_Servicios.Show
            frm_VTA_Servicios.SetFocus
        Case "007"
            Unload Me
            'objVenta.LimpiaProyecto
            objVenta.LimpiaServicio
            objVenta.LimpiaConvenio
            LimpiarSiSalgodeGuia
            frmPedido.grdPedido.Rebind
            objVenta.ptmModalidad = Cobranza_VtaCred
            objVenta.CodigoTipoVenta = Cobranza_VtaCred
            objVenta.PctBeneficiario = 0
            frmPedido.lblModalidad.Caption = objVenta.NombreTipoVenta
            'frmPedido.Label6.Visible = True
            'frmPedido.lblTotal.Visible = True
            frmPedido.Label4.Visible = False
            frmPedido.lblPctCopago.Visible = False
            frmPedido.Label8.Visible = False
            frmPedido.lblcopago.Visible = False
            frm_VTA_RecetarioM.pstrFlgRM = ""
            mdiPrincipal.cmdGrabaVenta.Enabled = True
            frm_VTA_Cobranza.Show
            frm_VTA_Cobranza.SetFocus
        Case "008"
'26/08/07 comentado por pHerrera, esta no es modalidad, pasa al boton en
' el formulario principal
'            Unload Me
'            frmPedido.grdPedido.Rebind
'            objVenta.LimpiaServicio
'            objVenta.LimpiaConvenio
'            LimpiarSiSalgodeGuia
'            frmPedido.grdPedido.Rebind
'            objVenta.ptmModalidad = Cotizaciones
'            objVenta.CodigoTipoVenta = Cotizaciones
'            objVenta.PctBeneficiario = 0
'            frmPedido.lblModalidad.Caption = objVenta.NombreTipoVenta
'            'frmPedido.Label6.Visible = True
'            'frmPedido.lblTotal.Visible = True
'            frmPedido.Label4.Visible = False
'            frmPedido.lblPctCopago.Visible = False
'            frmPedido.Label8.Visible = False
'            frm_VTA_RecetarioM.pstrFlgRM = ""
'            mdiPrincipal.cmdGrabaVenta.Enabled = True
'            frmPedido.lblcopago.Visible = False
'            frm_DLV_Pedido.blnActivaPed = False
'            frm_VTA_Cotizacion.Show vbModal
        Case "009"
            Unload Me
            objVenta.LimpiaServicio
            objVenta.LimpiaConvenio
            LimpiarSiSalgodeGuia
            frmPedido.grdPedido.Rebind
            objVenta.ptmModalidad = Venta_Regular
            objVenta.CodigoTipoVenta = FactServicios
            objVenta.PctBeneficiario = 0
            frmPedido.lblModalidad.Caption = objVenta.NombreTipoVenta
            'frmPedido.Label6.Visible = True
            'frmPedido.lblTotal.Visible = True
            frmPedido.Label4.Visible = False
            frmPedido.lblPctCopago.Visible = False
            frmPedido.Label8.Visible = False
            frmPedido.lblcopago.Visible = False
            frm_VTA_RecetarioM.pstrFlgRM = ""
            mdiPrincipal.cmdGrabaVenta.Enabled = True
            frm_VTA_FacServPrestados.Show
            frm_VTA_FacServPrestados.SetFocus
        Case "010" 'GUIA
            Unload Me
            objVenta.LimpiaServicio
            objVenta.LimpiaConvenio
            ptmTipoPrecio = Regular
            frmPedido.grdPedido.Rebind
            objVenta.ptmModalidad = Guias_Remision
            objVenta.CodigoTipoVenta = Guias_Remision
            objVenta.PctBeneficiario = 0
            frmPedido.lblModalidad.Caption = objVenta.NombreTipoVenta
            'frmPedido.Label6.Visible = True
            'frmPedido.lblTotal.Visible = True
            frmPedido.Label4.Visible = False
            frmPedido.lblPctCopago.Visible = False
            frmPedido.Label8.Visible = False
            frmPedido.lblcopago.Visible = False
            frmPedido.lblSiguiente.Caption = objUsuario.TipoDocGuia & " - " & objDocumento.ListaNumeroDisponible(objUsuario.CodigoEmpresa, objUsuario.NombrePC, objUsuario.TipoDocGuia)
            frm_VTA_RecetarioM.pstrFlgRM = ""
            mdiPrincipal.cmdGrabaVenta.Enabled = False
            frm_VTA_Busqueda.Datos Format(ptmTipoPrecio, "000")
            frm_VTA_Busqueda.SetFocus
            frm_VTA_Busqueda.txtBuscar.SetFocus

            'frm_VTA_GuiaRemision.Show
        Case "017"
            Unload Me
            objVenta.LimpiaServicio
            objVenta.LimpiaConvenio
            LimpiarSiSalgodeGuia
            frmPedido.grdPedido.Rebind
            objVenta.CodigoTipoVenta = Recetario
            objVenta.ptmModalidad = Recetario
            'ptmTipoPrecio = Reg_Mag
            objVenta.PctBeneficiario = 0
            'frmPedido.Label6.Visible = True
            'frmPedido.lblTotal.Visible = True
            frmPedido.lblModalidad.Caption = objVenta.NombreTipoVenta
            frmPedido.Label4.Visible = False
            frmPedido.lblPctCopago.Visible = False
            frmPedido.Label8.Visible = False
            frmPedido.lblcopago.Visible = False
'            Dim i As Integer
'            If frmPedido.grdPedido.ApproxCount > 0 Then
'                For i = 0 To objVenta.Producto.UpperBound(1)
'                    If objVenta.Producto(i, 7) <> objVenta.ptmModalidad Then
'                        frmPedido.grdPedido.Delete
'                        frmPedido.Cal_Promo
'                        frmPedido.Cal_Montos
'                        frmPedido.Refresh
'                    End If
'                Next i
'            End If
            
            mdiPrincipal.cmdGrabaVenta.Enabled = True
            frm_VTA_RecetarioM.pstrFlgRM = "1"
            frm_VTA_RecetarioM.Show
            frm_VTA_RecetarioM.SetFocus
        Case "019"
            Unload Me
            objVenta.LimpiaServicio
            objVenta.LimpiaConvenio
            LimpiarSiSalgodeGuia
            frmPedido.grdPedido.Rebind
            If objVenta.ptmModalidad = Venta_Convenio Then objVenta.CodigoConvenio = ""
            objVenta.ptmModalidad = Cajero_Corresponsal '  Venta_Regular
            ptmTipoPrecio = Regular
            objVenta.CodigoTipoVenta = Cajero_Corresponsal '  Venta_Regular
            objVenta.PctBeneficiario = 0
            frmPedido.lblModalidad.Caption = objVenta.NombreTipoVenta
            frmPedido.Label4.Visible = False
            frmPedido.lblPctCopago.Visible = False
            frmPedido.Label8.Visible = False
            frmPedido.lblcopago.Visible = False
            frmPedido.lblSiguiente.Caption = objUsuario.TipDocDefault & " - " & objDocumento.ListaNumeroDisponible(objUsuario.CodigoEmpresa, objUsuario.NombrePC, objUsuario.TipDocDefault)
            frmPedido.Cal_Montos
            mdiPrincipal.cmdGrabaVenta.Enabled = True
            frm_VTA_RecetarioM.pstrFlgRM = ""
            Select Case objUsuario.TipoMaquina
                Case objUsuario.TipoMaquinaAdmin
                    frm_VTA_Busqueda.Datos Format(ptmTipoPrecio, "000")
                Case objUsuario.TipoMaquinaCajero
                    frm_VTA_Busqueda.Datos Format(ptmTipoPrecio, "000")
                Case objUsuario.TipoMaquinaCabina
                Case objUsuario.TipoMaquinaRuteo
            End Select
            objVenta.CodigoTipoVenta = Cajero_Corresponsal
            frmCajeroExpress.Show
    End Select
    
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, "Error"
End Sub

Private Sub grdModalidad_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn
                 grdModalidad_DblClick
        Case vbKeySpace
                 grdModalidad_DblClick
    End Select
End Sub

Private Sub SeteaGrilla()
    
    Dim i%
    For i = 0 To grdModalidad.Columns.Count - 1
        grdModalidad.Columns(i).Visible = False
        grdModalidad.Columns(i).WrapText = True
    Next i
    
    grdModalidad.RowHeight = grdModalidad.RowHeight
    
    grdModalidad.Columns(2).FetchStyle = True
    
    grdModalidad.Columns(0).Visible = False
    grdModalidad.Columns(1).Visible = True
    grdModalidad.Columns(2).Visible = True
    grdModalidad.Style.VerticalAlignment = dbgVertCenter
    grdModalidad.Columns(2).Alignment = dbgCenter
    grdModalidad.Columns(2).ButtonText = True
    grdModalidad.HeadLines = 0
    
    grdModalidad.Columns(0).AllowFocus = False
    grdModalidad.Columns(1).AllowFocus = False
    'grdModalidad.Styles(5).BackColor = &HFFD7D7
    grdModalidad.Styles(5).Font.Bold = True
    
    'psub_Grilla_Traslate grdModalidad, "FLG_ACTIVO", "1", ilsImagenes.ListImages(1).Picture
End Sub

Sub LimpiaConvenio()
    objVenta.LimpiaConvenio
End Sub

Private Sub LimpiarSiSalgodeGuia()
    If objVenta.ptmModalidad = Guias_Remision Then
          ''  frmPedido.psub_BeginArry
            mdiPrincipal.subNuevo
    End If
End Sub

