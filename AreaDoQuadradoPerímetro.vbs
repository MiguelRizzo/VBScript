Dim n1, area, perimetro, resp

call entrada_numero
sub entrada_numero()

n1 = (inputbox("Digite Quantos cent�metros tem o lado do quadrado","Area e Perimetro"))

perimetro = (n1*4)
area = (n1*n1)

msgbox("O Per�metro �: "&perimetro&"" + vbnewline&_ 
		"A �rea �: "&area&""), vdInformation + vbOKOnly

call continuar
end sub

sub continuar()
resp=msgbox("Deseja fazer novamente?", vbquestion + vbyesno)
if resp=vbyes then
	call entrada_numero
else 
	wscript.quit
end if
end sub