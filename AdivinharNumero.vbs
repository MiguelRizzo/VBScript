dim numcerto, numescolhido, tentativas

call contador_tentativas
sub contador_tentativas()
tentativas = 5
randomize(second(time))
numcerto = int(rnd * 50) + 1
call escolha_numero
end sub

sub escolha_numero()

numcerto = numcerto
numescolhido = cdbl(inputbox("Escolha um nunero de 1 a 50", "Cinco tentativas"))
tentativas = tentativas - 1

if (numescolhido = numcerto) then
 MsgBox("Voce acertou!!! Voces escolheu "& numescolhido &" e numero certo era " & numcerto &"")
    call continuar

elseif tentativas = 0 then
 MsgBox("Voce esgotou suas tentativas. O numero certo era " & numcerto & "")
    call continuar

elseif (numescolhido > numcerto) then
 MsgBox("Voce errou. Coloque um numero mais baixo!" + vbNewLine &_
         "Aindam restam: " & tentativas & " tentativas")
 call escolha_numero

else
 MsgBox("Voce errou. Coloque um numero mais alto!" + vbNewLine &_
         "Aindam restam: " & tentativas & " tentativas")
 call escolha_numero

 end if

 call continuar
end sub

sub continuar()
resp=MsgBox("Deseja Realizar novo cálculo?", vbQuestion + vbYesNo, "Calcular Area e Perímetro")

if resp=vbyes Then
   call contador_tentativas
else
  wscript.quit
  end if
  end sub