dim nome_jogador, ataque_jogador, selecao_ataque, resp, n
dim selecao_ataque_cpu, ataque_cpu, resultado

call inicio

sub inicio()
nome_jogador=inputbox("Digite o nome do jogador: ","JOKENPO")
call jogar
end sub

sub jogar()
selecao_ataque=cint(inputbox("[1] Papel" + vbnewline & _
                             "[2] Pedra" + vbnewline & _
					         "[3] Tesoura" + vbnewline & _
							 "[4] Encerar Jogo",""& nome_jogador &" Escolha seu Ataque"))

select case selecao_ataque
            case 1:
			     ataque_jogador=""& nome_jogador &" escolheu Papel"
			case 2:
			     ataque_jogador=""& nome_jogador &" escolheu Pedra"
			case 3:
			     ataque_jogador=""& nome_jogador &" escolheu Tesoura"
		    case 4:
			     resp=msgbox("Deseja encerrar o jogo?",vbquestion+vbyesno,"ATENÇÃO")
				 if resp=vbyes then
				    wscript.quit
				 else 
				    call jogar
				 end if
			case else 
			     msgbox("Opção de ataque inválida!"),vbexclamation+vbokonly,"ATENÇÃO"
				 call jogar 
end select

randomize(second(time))  
selecao_ataque_cpu=int(rnd * 3) + 1

select case selecao_ataque_cpu
            case 1:
			     ataque_CPU="O Computador escolheu Papel"
			case 2:
			     ataque_CPU="O Computador escolheu Pedra"
			case 3:
			     ataque_CPU="O Computador escolheu Tesoura"
end select

if selecao_ataque = selecao_ataque_cpu then
   resultado="Houve um Empate!"
elseif selecao_ataque = 1 and selecao_ataque_cpu = 2 then
   resultado="Você venceu!"
elseif selecao_ataque = 1 and selecao_ataque_cpu = 3 then
   resultado="Você Perdeu!"
elseif selecao_ataque = 2 and selecao_ataque_cpu = 1 then
   resultado="Você Perdeu!"
elseif selecao_ataque = 2 and selecao_ataque_cpu = 3 then
   resultado="Você Ganhou!"
elseif selecao_ataque = 3 and selecao_ataque_cpu = 1 then
   resultado="Você Ganhou!"
elseif selecao_ataque = 3 and selecao_ataque_cpu = 2 then
   resultado="Você Perdeu!"
end if


msgbox(""& ataque_jogador &"" + vbnewline & _
       ""& ataque_CPU &"" + vbnewline & _
	   ""& resultado &""),vbinformation + vbokonly,"RESULTADO DO COMBATE - JO-KEN-PO"
call jogar
end sub