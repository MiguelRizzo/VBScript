dim palavras(30), i, resp, pontos, acertos, pular, jogarnovamente, audio

call valores
sub valores()
acertos = 0
pontos = 0
pular = 1
call carregar_voz
end sub

sub carregar_voz()
set audio=createobject("SAPI.SPVOICE")
audio.volume=100
audio.rate=-1

call carregar_palavras
end sub

sub carregar_palavras()
palavras(1)="Árvore"
palavras(2)="Televisão"
palavras(3)="Sapatol"
palavras(4)="Anel"
palavras(5)="Porta"
palavras(6)="Cama"
palavras(7)="Chão"
palavras(8)="Aplicativo"
palavras(9)="Cachorro"
palavras(10)="Cabelo"
palavras(11)="Amarelo"
palavras(12)="Camiseta"
palavras(13)="Escova"
palavras(14)="Espelho"
palavras(15)="Relógio"
palavras(16)="Cadeado"
palavras(17)="Lousa"
palavras(18)="Gato"
palavras(19)="Mochila"
palavras(20)="Armário"
palavras(21)="Cérebro"
palavras(22)="Caderno"
palavras(23)="Amigo"
palavras(24)="Macaco"
palavras(25)="Teclado"
palavras(26)="Colher"
palavras(27)="Igreja"
palavras(28)="Carro"
palavras(29)="Óculos"
palavras(30)="Papel"
call randomizar
end sub

sub randomizar()
randomize(second(time))
i=int(rnd * 30) + 1

call soletrando
end sub

sub falar_audio
audio.speak(palavras(i))
end sub

sub soletrando()

resp = (inputbox("Digite a palavra ouvida" + vbNewLine & _
"[O] Ouvir a palavra" + vbNewLine & _
"[R] Repetir a Palavra" + vbNewLine & _
"[P] Pular a palavra (1x)", "Soletrando"))

if (resp = "O") then
call falar_audio
call soletrando

elseif (resp = "R") then
	call falar_audio
	call soletrando

elseif (resp = "P") and (pular = 1) then
pular = 0
	call randomizar

elseif (resp = "P") and (pular < 1) then
MsgBox("Você já utilizou esta ajuda 1 vez")
	call soletrando

elseif (acertos = 30) then
jogarnovamente = MsgBox("Voce Ganhou!!!" + vbNewLine & _
"Acertos: " & acertos & + vbNewLine & _
"Pontuação: " & pontos & + vbNewLine & _
"Deseja jogar novamente? ", vbQuestion + vbYesNo, "")
if jogarnovamente=vbyes Then
call valores
	else
	wscript.quit
	end if

elseif (resp = palavras(i)) then
acertos = acertos + 1
pontos = pontos + 1000
MsgBox("Voce acertou!!!" + vbNewLine & _
"Palavra: " & palavras(i) & + vbNewLine & _
"Acertos: " & acertos & + vbNewLine & _
"Pontos: " & pontos & "")
	call randomizar





else
jogarnovamente = MsgBox("Voce errou!!!" + vbNewLine & _
"Acertos: " & acertos & + vbNewLine & _
"Pontos: " & pontos & + vbNewLine & _
"Deseja jogar novamente? ", vbQuestion + vbYesNo, "")
if jogarnovamente=vbyes Then
	call valores
	else
	wscript.quit
	end if

	end if

end sub