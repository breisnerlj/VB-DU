VERSION 5.00
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "oradc.ocx"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form frm_VTA_TipoMaquina 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Máquina"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin ORADCLibCtl.ORADC oradcTipoMaquina 
      Height          =   375
      Left            =   1920
      Top             =   1560
      Visible         =   0   'False
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   207
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.24
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
   Begin TrueDBGrid70.TDBGrid grdTipoMaquina 
      Bindings        =   "frm_VTA_TipoMaquina.frx":0000
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6588
      _LayoutType     =   4
      _RowHeight      =   31
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Codigo"
      Columns(0).DataField=   "COD"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Descripción"
      Columns(1).DataField=   "DES"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   528
      Columns(2)._MaxComboItems=   5
      Columns(2).ValueItems(0)._DefaultItem=   0
      Columns(2).ValueItems(0).Value=   "002"
      Columns(2).ValueItems(0).Value.vt=   8
      Columns(2).ValueItems(0).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(2).ValueItems(0).DisplayValue(0)=   "bHQAADYOAABCTTYOAAAAAAAANgAAACgAAAAlAAAAIAAAAAEAGAAAAAAAAA4AAAAAAAAAAAAAAAAA"
      Columns(2).ValueItems(0).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(2)=   "///////////////9/f3u6+ff2NLs6+b/////////////////////////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(3)=   "//8A////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(4)=   "//////DuzaqdjGtTZzcZYhwAbRIApoF1/f38////////////////////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(5)=   "AP////////////////////////////////////////////////////////////////////3c1Jl2"
      Columns(2).ValueItems(0).DisplayValue(6)=   "YUxgSjtrVGcgAG8jADObghO6wUgsC6yBbf///////////////////////////////////////wD/"
      Columns(2).ValueItems(0).DisplayValue(7)=   "///////////////////////////////////////////////////////96eXFpZdndmETi4AA3eQA"
      Columns(2).ValueItems(0).DisplayValue(8)=   "//8iu6x6KAB5KAAkwK8A//8A8PBIcV766+f///////////////////////////////////8A////"
      Columns(2).ValueItems(0).DisplayValue(9)=   "//////////////////////////////////////////Xz07apmY5+SXVhEKygAO35AP//APj/AP//"
      Columns(2).ValueItems(0).DisplayValue(10)=   "J7SjeSYAeCcAJrupAPv/AP//AK+ryr2y////////////////////////////////////AP//////"
      Columns(2).ValueItems(0).DisplayValue(11)=   "/////////////////////////////+zo5bmci3JrVSSJfgDDwAD5/wD//wD1/wDy/wDx/wD//yW3"
      Columns(2).ValueItems(0).DisplayValue(12)=   "qHomAHgnACa7qQD8/wD0/wD2/19mUf///////////////////////////////////wD/////////"
      Columns(2).ValueItems(0).DisplayValue(13)=   "///////////////9/f3f1tOHaVVhIABGQTAA5+cA//8A+f8A8/8A8f8A8f8A8f8A8f8A//8jvK56"
      Columns(2).ValueItems(0).DisplayValue(14)=   "JwB4JwAmu6kA/P8A8P8A//8ShXf/29b///////////////////////////////8A////////////"
      Columns(2).ValueItems(0).DisplayValue(15)=   "////////2tHHbT8nQAAAXx8AeCcAWV5HAff7APT/APH/APH/APL/APL/APL/APH/AP//I7yveScA"
      Columns(2).ValueItems(0).DisplayValue(16)=   "eCcAJrupAPz/APH/APT/APX/q25V////////////////////////////////AP//////////////"
      Columns(2).ValueItems(0).DisplayValue(17)=   "/97TymczEWMiAHU1AHQ1AXkqAFhbRQH0+AD0/wDx/wDy/wDy/wDy/wDy/wDx/wD+/yDAtXkmAHgn"
      Columns(2).ValueItems(0).DisplayValue(18)=   "ACa7qQD8/wDx/wDy/wD7/2CGd/7l4P///////////////////////////wD///////////////+k"
      Columns(2).ValueItems(0).DisplayValue(19)=   "gnJgHgB0MwBzNAFyNAF5KgBbWUIC8vYA9P8A8f8A8v8A8v8A8v8A8v8A8f8A/f8cyb55JwB4JwAm"
      Columns(2).ValueItems(0).DisplayValue(20)=   "u6kA/P8A8f8A8f8A+v8Yr6zRuqz///////////////////////////8A/////////////fz6mXBQ"
      Columns(2).ValueItems(0).DisplayValue(21)=   "eT4OcDAAcjMAcjQBdyoAXVY7Be3xAPX/APH/APL/APL/APL/APL/APH/AP3/GMzFeCcBeyIAKLek"
      Columns(2).ValueItems(0).DisplayValue(22)=   "APz/APH/APH/APf/ANXXiJWI//z7////////////////////////AP///////////////6+Lb4xa"
      Columns(2).ValueItems(0).DisplayValue(23)=   "MXAwAHEyAHM0AXYsAGVJKAvk4wD2/wDx/wDy/wDy/wDy/wDy/wDx/wD//yqxoIIXAHMvAB3JvQD6"
      Columns(2).ValueItems(0).DisplayValue(24)=   "/wDx/wDx/wDz/wD3/zJ+cPXn5P///////////////////////wD////////////////Sv6+cb0x6"
      Columns(2).ValueItems(0).DisplayValue(25)=   "Pg1vLwBzNAFzMQB4KQAZzsUA+v8A8f8A8f8A8v8A8v8A8f8A8v8A//9JeFpObVEO3d0A9/8A8v8A"
      Columns(2).ValueItems(0).DisplayValue(26)=   "8f8A8v8A8f8A//8AqqDHraL///////////////////////8A/////////////////fz6nnVRkmA5"
      Columns(2).ValueItems(0).DisplayValue(27)=   "ayoAczQBczIAdysCSHZiAP//APf/APH/APH/APH/APH/AP//IcK1L56RAP//APn/APL/APH/APH/"
      Columns(2).ValueItems(0).DisplayValue(28)=   "APL/APH/APP/AP//YVc+////////////////////////AP///////////////////8GlkJltSHEy"
      Columns(2).ValueItems(0).DisplayValue(29)=   "AXEzAHcsAFVkSj2IcTuOeQPx9wD//wD7/wD+/wD//x3GvUR/aAvh5QD3/wDx/wDx/wDw/wDz/wDw"
      Columns(2).ValueItems(0).DisplayValue(30)=   "/wDw/wDw/wD//xCmn+7Gvf///////////////////wD////////////////////l2tKWZkCBSiFt"
      Columns(2).ValueItems(0).DisplayValue(31)=   "LAB0MQBzMQQM5uEtpZppPxI5lH8O2twZx8NIeV5JcVYM3uIA+f8A8/8A+v8A//8A//8A//8A//8A"
      Columns(2).ValueItems(0).DisplayValue(32)=   "//8A//8A//8KqqC5hnP///////////////////8A////////////////////////lmZBm21IcjMB"
      Columns(2).ValueItems(0).DisplayValue(33)=   "cTIAdi0AgRkAGcjFCujnSXZcaEMXaEMWNpODAP//AP//APn/BPLyE9HUNpWGdykAlAAAcDQPO415"
      Columns(2).ValueItems(0).DisplayValue(34)=   "SnVXeyQAhAcAXjQf9fTx////////////////AP///////////////////////9jHuoVPJH1EFnAw"
      Columns(2).ValueItems(0).DisplayValue(35)=   "AHI0AXUuAH4dADuOdwnl6AD//wHx+TCjk1pbNzKcikGCa1pbPHE0Bn4fAHwiAFpcORHa1gD1/wD0"
      Columns(2).ValueItems(0).DisplayValue(36)=   "/hHX1GNQK18TAMCxpv///////////////wD////////////////////////18e2YakWIUiduLQBx"
      Columns(2).ValueItems(0).DisplayValue(37)=   "MwB5JwBYXD0A9v8A+v8A8/8A9v8A//82mImGEwB3KwB4KgB1LwB0MQB4JwAE7PMA//8A8/8A9P8A"
      Columns(2).ValueItems(0).DisplayValue(38)=   "//8R2ttnDwCObF////////////////8A////////////////////////////vaCHlWU/bi0AcTMA"
      Columns(2).ValueItems(0).DisplayValue(39)=   "fSEAMKOTAP//APD/APH/APH/APb/APP+fh8AczIAcjQBcjQBeiYAUmZFAP//APL/APH/APH/APP/"
      Columns(2).ValueItems(0).DisplayValue(40)=   "AP//WE4rZicJ6+jk////////////AP///////////////////////////9rLv6N7Wng8C28vAH0i"
      Columns(2).ValueItems(0).DisplayValue(41)=   "AC+klQD//wDw/wDy/wDx/wD3/wHv+X0gAHMyAHM0AXI0AXkoAFVjQQD//wDy/wDx/wDx/wD0/wD/"
      Columns(2).ValueItems(0).DisplayValue(42)=   "/11NKWQfAOjl4P///////////wD////////////////////////////28e+of2GRYDhrKwB5JwBW"
      Columns(2).ValueItems(0).DisplayValue(43)=   "ZEEA//8A+P8A8v8A8v8A//8sqZ6BGgByNQFzNAFyNAF0MAB3KgMC7/YA//8A8v8A8v8A//8H5etu"
      Columns(2).ValueItems(0).DisplayValue(44)=   "GQB7Szf29vP///////////8A////////////////////////////////s5J3o3hXczUDcS8AeSgA"
      Columns(2).ValueItems(0).DisplayValue(45)=   "QYVvA/T4AP7/AP79G8jDdC8AdTAAcjQBczQBczQBcjQAeCgAW1k0DOLhAP//AP7+C9/faE0jZhcA"
      Columns(2).ValueItems(0).DisplayValue(46)=   "zLyt////////////////AP////////////////////////////////7//4JKHo5bNG0uAHMyAH8d"
      Columns(2).ValueItems(0).DisplayValue(47)=   "AHcqAFFoTmo/FoAaAHUuAHI0AXM0AXM0AXM0AXIzAG8xAHwoAIgkAFZmR2FLJIw1B51hPeTb0f//"
      Columns(2).ValueItems(0).DisplayValue(48)=   "/////////////////wD///////////////////////////////////+ab0uRYDhyMwJxMgByNQF0"
      Columns(2).ValueItems(0).DisplayValue(49)=   "MAB6JwB3LABzMwByNAFzNAFyMwBxMgBwMABxMQB9QRJ0NgdgGQCFPBDHn4rg0sb9/v//////////"
      Columns(2).ValueItems(0).DisplayValue(50)=   "//////////////8A////////////////////////////////////4dTLi1cufkMWcDAAczQBcjQB"
      Columns(2).ValueItems(0).DisplayValue(51)=   "cjQBcjQBcjQBcTEAcTIAcjMAdjgFdzwMeT0OdDgGmXBP5dzV/Pv7////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(52)=   "////////////AP////////////////////////////////////fz8qJ5WYZPJG0sAHI0AHIzAHIz"
      Columns(2).ValueItems(0).DisplayValue(53)=   "AHEyAHExAHY5CHc6B28wA3Y6CpNkPr2ijfDq5v//////////////////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(54)=   "/////////wD///////////////////////////////////////++pIykfFt0NgdrKQBtLQB0NQN4"
      Columns(2).ValueItems(0).DisplayValue(55)=   "Owt3OwlwMACASx2siXXWxbn9/fz/////////////////////////////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(56)=   "//////8A////////////////////////////////////////7+jjs5N3r41yj143hEwfdDUCbi8A"
      Columns(2).ValueItems(0).DisplayValue(57)=   "jF02x7Ke9/Xy////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(58)=   "////AP////////////////////////////////////////////r49raVe5xvS4VQL6R/YfLu6///"
      Columns(2).ValueItems(0).DisplayValue(59)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(60)=   "/wD/////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(61)=   "//////////////////////////////////////////////////////////////////////////8A"
      Columns(2).ValueItems(0).DisplayValue(62)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(0).DisplayValue(63)=   "////////////////////////////////////////////////////////////////////////AA=="
      Columns(2).ValueItems(0).DisplayValue.vt=   9
      Columns(2).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems(1)._DefaultItem=   0
      Columns(2).ValueItems(1).Value=   "003"
      Columns(2).ValueItems(1).Value.vt=   8
      Columns(2).ValueItems(1).DisplayValue.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      Columns(2).ValueItems(1).DisplayValue(0)=   "bHQAADYOAABCTTYOAAAAAAAANgAAACgAAAAlAAAAIAAAAAEAGAAAAAAAAA4AAAAAAAAAAAAAAAAA"
      Columns(2).ValueItems(1).DisplayValue(1)=   "AAAAAAD/////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(1).DisplayValue(2)=   "///////////////9/f3u6+je2NLt6uf/////////////////////////////////////////////"
      Columns(2).ValueItems(1).DisplayValue(3)=   "//8A////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(1).DisplayValue(4)=   "/////P39xbitjWxUZzcbWiYBXC8ToIl7/Pz9////////////////////////////////////////"
      Columns(2).ValueItems(1).DisplayValue(5)=   "AP////////////////////////////////////////////////////////////////////Lw7pV7"
      Columns(2).ValueItems(1).DisplayValue(6)=   "amEyKV8nAWUlAGsrAHAvAGopAFIZAKWPgv///////////////////////////////////////wD/"
      Columns(2).ValueItems(1).DisplayValue(7)=   "/////////////////////////////////////////////////////f3z9PGpx8F+SChSDwBeHABu"
      Columns(2).ValueItems(1).DisplayValue(8)=   "LwBzMgBzNABzNAFzNAF0NQFpKABkNRn28/L///////////////////////////////////8A////"
      Columns(2).ValueItems(1).DisplayValue(9)=   "//////////////////////////////////////////7+y8O5c87NDuLqAPz/NnBXaUAXdjIAdTAA"
      Columns(2).ValueItems(1).DisplayValue(10)=   "cjQBczQBcjQBcjQBczQBczQATRAAyby0////////////////////////////////////AP//////"
      Columns(2).ValueItems(1).DisplayValue(11)=   "/////////////////////////////+3q57ShknxWQ1wkBENOMzKXgxXSzgD//yS1rUOBaHwjAHQx"
      Columns(2).ValueItems(1).DisplayValue(12)=   "AHI0AXI0AXM0AXM0AXQ1AV8iAG1HK////////////////////////////////////wD/////////"
      Columns(2).ValueItems(1).DisplayValue(13)=   "///////////////9/f3f19OHaFVbJwJaHQBoJwByMQB4LAB8JQB3KgBjTSJGfWYA9PxAh3F5JgB0"
      Columns(2).ValueItems(1).DisplayValue(14)=   "MQBzMwByNAFzNAFzNAFyMgBIEQD39vf///////////////////////////////8A////////////"
      Columns(2).ValueItems(1).DisplayValue(15)=   "////////2tDHbEAnQQEAYB8AcDAAdDQAczQAczQBcjQAcjQAczEAdywAiBEATXFVHr+5VmNCcDUE"
      Columns(2).ValueItems(1).DisplayValue(16)=   "dDAAczMAczQBczQBdDMAXR4Am4V0////////////////////////////////AP//////////////"
      Columns(2).ValueItems(1).DisplayValue(17)=   "/97TymczEWIiAHU1AHQ1AXM0AXI0AXI0AXM0AXM0AXM0AXM0AXI1AXskAEOCakxxUkp1WU1tTmRK"
      Columns(2).ValueItems(1).DisplayValue(18)=   "IHgrAHI0AHM0AXM0AWspAHNONvf29f///////////////////////////wD///////////////+j"
      Columns(2).ValueItems(1).DisplayValue(19)=   "gXJiHQB0NABzNAFyNAFzNAB1MAB0MgByMwByNAFyNAFzNAFyNAF9IQA/iXQ6kHw5kn84lIJFfGNk"
      Columns(2).ValueItems(1).DisplayValue(20)=   "TCV2LAByNABzNAFyMwBbJgbKvbX///////////////////////////8A/////////////Pv6mW5P"
      Columns(2).ValueItems(1).DisplayValue(21)=   "eT0OcDAAcjMAczQAdTAAaz8abzYLeCkAeCoAdi4AczMAcjQAfiAAOpF/P4RuAfb9NZeGPYpzQIVs"
      Columns(2).ValueItems(1).DisplayValue(22)=   "aj8YdyoAcjQBdDUBWxwAl3ho////////////////////////////AP///////////////6+Lb41Y"
      Columns(2).ValueItems(1).DisplayValue(23)=   "MHAxAHEyAHMyAHYsAB3CuFJtSkp1WlpbPWpBFngpAHcqAIEbADOaikGEbAjq7Aju7UKAaj6IcjWX"
      Columns(2).ValueItems(1).DisplayValue(24)=   "hXQwBXYsAHM1AWsrAFopCfHv7P///////////////////////wD////////////////Tv66gbUeH"
      Columns(2).ValueItems(1).DisplayValue(25)=   "MQB0KAB7JQBQakoN3N9CgWkznIxGe2AP1to/hW9hTyZ3KgIuoZVBhGwH6OwA//8hvrREgGg/iXIp"
      Columns(2).ValueItems(1).DisplayValue(26)=   "rKF6KAB1LwB0NAFMDgDCtq7///////////////////////8A//////////////////v6UbetL7iw"
      Columns(2).ValueItems(1).DisplayValue(27)=   "Ul5FaEIqaUMUYVAqSnNZNJiHGcbETHFTGsXCFM7McTUaLKWbQYFqBevxAPX/AP//NZeGPIx4NZiH"
      Columns(2).ValueItems(1).DisplayValue(28)=   "KaygeCoAdTIAZSYAa0Ek////////////////////////AP///////////////////8mehHuGazeU"
      Columns(2).ValueItems(1).DisplayValue(29)=   "gSyjkzSZh0V9aV5WM3A2B1xXP2w+DCavpS6kmYoNACWxqkKCawTu9QD0/wD1/wD2/kp2WjWXhxzC"
      Columns(2).ValueItems(1).DisplayValue(30)=   "vzOcjIIZAHIxAFIZAOPd2v///////////////////wD////////////////////l2tGdYDeJPhRw"
      Columns(2).ValueItems(1).DisplayValue(31)=   "KgJhTyhNcFEtpJUP19kwoJh0LwB4KgFsPA6MBwAevLpCgWkD8PkA9P8A8f8A+v8L4ORYXTc7jHgA"
      Columns(2).ValueItems(1).DisplayValue(32)=   "9/9QaU6CHgBWGACtmYv///////////////////8A////////////////////////lWZBmm1IcjIA"
      Columns(2).ValueItems(1).DisplayValue(33)=   "dC0AeigAgB0AhRMAWV04FszOJq6lTHJUYFAlFszLSXdZAPb/APP/APH/APH/AP//G8TARX5lTnFT"
      Columns(2).ValueItems(1).DisplayValue(34)=   "AP//S3VccyMAYDIe9PTx////////////////AP///////////////////////9nHuoRPJH1EFnAw"
      Columns(2).ValueItems(1).DisplayValue(35)=   "AHI0AXI1AXI0AHkmAH8eAEOBaER+ZUl2Wl5UOExyUwD6/wDy/wDy/wDx/wDz/wHz+jiTgS2mmjWZ"
      Columns(2).ValueItems(1).DisplayValue(36)=   "hwD//2pEHlkbAMCxpv///////////////wD////////////////////////18e2YakWIUiduLQBy"
      Columns(2).ValueItems(1).DisplayValue(37)=   "MwBzNAFzNAFyNAB4KgBPbE4F7fIbyME8i34R1NYA/P8A9P8A8f8A8f8A8f8A//8S19ZLdFkcv7tN"
      Columns(2).ValueItems(1).DisplayValue(38)=   "cVByOQheHQCMbmD///////////////8A////////////////////////////vaCHlWU/bi0AcTIA"
      Columns(2).ValueItems(1).DisplayValue(39)=   "czQBczQBcjQBczQAfSEAPYdzAP//AP//AP//Dd3iBeXvAPb/APX/AP3/Hb69P4dvN5SBYk8meiUA"
      Columns(2).ValueItems(1).DisplayValue(40)=   "dDIAbCwAYTAQ6+fk////////////AP///////////////////////////9rLv6N7Wng7C3AvAHM0"
      Columns(2).ValueItems(1).DisplayValue(41)=   "AXM0AXM0AXI0AXQxAHgoADqSgAD//xfLyCO5r0CGbwD7/wTx9Curnj2JdTiSf284B3omAHIzAHI0"
      Columns(2).ValueItems(1).DisplayValue(42)=   "AHAvAF0pBenk4P///////////wD////////////////////////////28e+of2CRYjtrKgBzNAFz"
      Columns(2).ValueItems(1).DisplayValue(43)=   "NAFzNAFzNAFyNAF1LwB/HgAms6cWyswhubQzmowL3+U3l4QspZcwoJNdVi91LQByNABxMQByMwBm"
      Columns(2).ValueItems(1).DisplayValue(44)=   "JgB5TTn29fP///////////8A////////////////////////////////s5J3o3hWczQDcDEAczQB"
      Columns(2).ValueItems(1).DisplayValue(45)=   "czQBczQBczQBcjQBdi8AdiwAE9bUGczECO7tN5OCHr24OJF9Q4BmdygAcjAAcTIAczUCdjcFYCED"
      Columns(2).ValueItems(1).DisplayValue(46)=   "zbut////////////////AP////////////////////////////////7//4JKHo5bNG0tAHM0AXM0"
      Columns(2).ValueItems(1).DisplayValue(47)=   "AXM0AXM0AXM0AXI0AXgqAGZEHRPW0Ud6YB6/u1ZiPy2fkIEhAHo6CnU3BnQ2BIJLIJdsSeXa0f//"
      Columns(2).ValueItems(1).DisplayValue(48)=   "/////////////////wD///////////////////////////////////+bb0uQYDhyMwJxMgBzNAFz"
      Columns(2).ValueItems(1).DisplayValue(49)=   "NAFzNAFzNAFzNAFzMwBzMQBrPhRqPhBbUSp1LAB7QxR2MgRgGgB/SB3Bp5Le08j9/v7/////////"
      Columns(2).ValueItems(1).DisplayValue(50)=   "//////////////8A////////////////////////////////////4dTLi1cufUQWcDAAczQBczQB"
      Columns(2).ValueItems(1).DisplayValue(51)=   "czQBczQBczQAcTIAcjIAdDAAdzUBezcFejwMdTYDmXBP5dzV/Pv6////////////////////////"
      Columns(2).ValueItems(1).DisplayValue(52)=   "////////////AP////////////////////////////////////fz8aJ6WYVPI20sAHI0AHIzAHIz"
      Columns(2).ValueItems(1).DisplayValue(53)=   "AHEyAHAxAHY5CHc6B3AwA3Y6CZNkP72ijfDq5///////////////////////////////////////"
      Columns(2).ValueItems(1).DisplayValue(54)=   "/////////wD///////////////////////////////////////++pI2kfFt0NgdrKQBtLQB0NQN4"
      Columns(2).ValueItems(1).DisplayValue(55)=   "Owt2OwlwMQCASR6riXXWxbj9/fz/////////////////////////////////////////////////"
      Columns(2).ValueItems(1).DisplayValue(56)=   "//////8A////////////////////////////////////////7+jjs5N3sI5ykF43g0sfdDUCbi8A"
      Columns(2).ValueItems(1).DisplayValue(57)=   "jV02x7Ke9/Tz////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(1).DisplayValue(58)=   "////AP////////////////////////////////////////////r39raWe5xvS4ZQL6R/YfLu6///"
      Columns(2).ValueItems(1).DisplayValue(59)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(1).DisplayValue(60)=   "/wD/////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(1).DisplayValue(61)=   "//////////////////////////////////////////////////////////////////////////8A"
      Columns(2).ValueItems(1).DisplayValue(62)=   "////////////////////////////////////////////////////////////////////////////"
      Columns(2).ValueItems(1).DisplayValue(63)=   "////////////////////////////////////////////////////////////////////////AA=="
      Columns(2).ValueItems(1).DisplayValue.vt=   9
      Columns(2).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(2).ValueItems.Count=   2
      Columns(2).Caption=   "COD_MENU"
      Columns(2).DataField=   "COD"
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
End
Attribute VB_Name = "frm_VTA_TipoMaquina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objMaquina As New clsMaquina
Dim odynR1 As oraDynaset
Dim odynClon As oraDynaset
Dim oraUsuario As oraDynaset
Private lblnVentana As Boolean



Private Sub Form_Load()
    lblnVentana = False
    
    Set oradcTipoMaquina.Recordset = objMaquina.ListaTipoMaquina
    Call SeteaGrilla
    
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Control
    Select Case KeyCode
        Case vbKeyReturn
            grdTipoMaquina_DblClick
        Case vbKeyEscape
            gclsOracle.Cerrar
            End
    End Select
Exit Sub
Control:
    MsgBox Err.Description, vbOKOnly + vbCritical, App.ProductName
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' If objVenta.CodigoTipoVenta < 1 Then Cancel = 1
End Sub

Private Sub grdTipoMaquina_DblClick()
Dim gvarError  As String
Dim strCodLocal As String
Dim strFlgDelivery As String

On Error GoTo handle

    If grdTipoMaquina.ApproxCount <= 0 Then Exit Sub
    
    
    If objMaquina.Valida Then
        Select Case grdTipoMaquina.Columns(0).Value
            Case "001", "002"
                strCodLocal = Mid(objMaquina.NombrePC, 4, 3)
                strFlgDelivery = "0"
            Case "003", "004", "005"
                strCodLocal = objUsuario.LocalDelivery
                strFlgDelivery = "1"
        End Select

    
    
        gvarError = objMaquina.Graba(objUsuario.CodigoEmpresa, _
                                    strCodLocal, objMaquina.NombrePC, _
                                    grdTipoMaquina.Columns(0).Value, _
                                     strFlgDelivery, objMaquina.NumIP, _
                                    "1", objUsuario.Codigo)
    
    
    
        If gvarError <> "" Then GoTo salir
    
        Set oraUsuario = objUsuario.Login(objUsuario.Codigo, objUsuario.Password)
        Unload Me
    Else
          GoTo salir
    End If
    
    
    Exit Sub
    
handle:
        MsgBox Err.Description, vbCritical + vbOKOnly, App.ProductName
salir:
        gclsOracle.Cerrar
        End
    
    
End Sub

Private Sub grdTipoMaquina_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn
                 grdTipoMaquina_DblClick
        Case vbKeySpace
                 grdTipoMaquina_DblClick
    End Select
End Sub

Private Sub SeteaGrilla()
    
    Dim i%
    For i = 0 To grdTipoMaquina.Columns.Count - 1
        grdTipoMaquina.Columns(i).Visible = False
        grdTipoMaquina.Columns(i).WrapText = True
    Next i
    
    grdTipoMaquina.RowHeight = grdTipoMaquina.RowHeight
    
    grdTipoMaquina.Columns(2).FetchStyle = True
    
    grdTipoMaquina.Columns(0).Visible = False
    grdTipoMaquina.Columns(1).Visible = True
    grdTipoMaquina.Columns(2).Visible = True
    grdTipoMaquina.Style.VerticalAlignment = dbgVertCenter
    grdTipoMaquina.Columns(2).Alignment = dbgCenter
    grdTipoMaquina.Columns(2).ButtonText = True
    grdTipoMaquina.HeadLines = 0
    
    grdTipoMaquina.Columns(0).AllowFocus = False
    grdTipoMaquina.Columns(1).AllowFocus = False
    'grdTipoMaquina.Styles(5).BackColor = &HFFD7D7
    grdTipoMaquina.Styles(5).Font.Bold = True
    
    'psub_Grilla_Traslate grdTipoMaquina, "FLG_ACTIVO", "1", ilsImagenes.ListImages(1).Picture
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



