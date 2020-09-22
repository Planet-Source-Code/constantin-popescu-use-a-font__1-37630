Add the file 'modFont.bas' to your app. and you will be able to 
use a font that is not installed on the system.

DON'T STOP THE APP. BY CLICKING END BUTTON IN VISUAL BASIC -
- THE FONT WILL REMAIN IN MEMORY!

PLEASE USE ONY THE Correct VERSION !

For one font 
Wrong >		Label1.FontName = UseFont("C:\fonts\font.ttf")
		Label2.FontName = UseFont("C:\fonts\font.ttf")

Correct > 	Dim fntFileName As string, fntName As String
		Private Sub Form_Load()
		fntFileName = "C:\fonts\font.ttf"
		fntName = UseFont(fntFileName)
		Label1.FontName = fntName 
		Label2.FontName = fntName 
		End Sub
		
		Private Sub Form_Unload(Cancel As Integer)
		RemoveFont (fntFileName)
		End Sub

For two or more fonts 

Wrong >		Label1.FontName = UseFont("C:\fonts\font01.ttf")
		Label2.FontName = UseFont("C:\fonts\font02.ttf")

Correct > 	Dim fntFileName01 As string, fntName01 As String
		Dim fntFileName02 As string, fntName02 As String
		Private Sub Form_Load()
		fntFileName01 = "C:\fonts\font01.ttf"
		fntFileName02 = "C:\fonts\font02.ttf"
		fntName01 = UseFont(fntFileName01)
		fntName01 = UseFont(fntFileName01)
		Label1.FontName = fntName01 
		Label2.FontName = fntName02 
		End Sub

		Private Sub Form_Unload(Cancel As Integer)
		RemoveFont (fntFileName01)
		RemoveFont (fntFileName02)
		End Sub

REMEMBER to remove the font(s) that you used otherwise the font(s) 
will temporary remain in your system and you will not be able to
move or delete this file(s) until you restart the computer.

Still if you don't remove the font, programs such as Word will 
recognize it in the font list. After restart the font will
dissapear from the list.

Remove only the font(s) that you added !