dim palavra, tentativa, audio, nivel1, nivel2, nivel3, nivel4, nome, acertos, resp

 call carregar_nome
 sub carregar_nome()
 nome=inputbox("DIGITE SEU NOME:" + vbNewLine)
 end sub

 call carregar_audio
 sub carregar_audio()
 set audio=createobject("SAPI.SPVOICE")
 audio.volume=100
 audio.rate = 0
 call carregar_nivel1
 end sub
 
 sub carregar_nivel1
 randomize(second(time))
 nivel1=int(rnd*5) + 1
 select case nivel1
 case 1:
 nivel1= "LIMÃO"
 case 2:
 nivel1= "BANANA"
 case 3:
 nivel1= "EGITO"
 case 4:
 nivel1= "NAVIO"
 case 5:
 nivel1= "JACARÉ"
 end select
 call carregar_tentativa1
 end sub
 
 sub carregar_tentativa1()
 audio.speak (""& nivel1 &"")
 tentativa=inputbox("DIGITE A PALAVRA OUVIDA" + vbNewLine + vbNewLine & _
					"NOME DO JOGADOR(A): " & nome &"" + vbNewLine + vbNewLine & _
					"[O] OUVIR NOVAMENTE A PALAVRA" + vbNewLine & _
					"[P] PULAR A PALAVRA UMA ÚNICA VEZ", "SOLETRANDO")
 if tentativa = nivel1 then
	msgbox("PARABÉNS, VOCÊ ACERTOU A PALAVRA!" + vbNewLine & _
		   "QUANTIDADE DE ACERTOS:" & acertos + 1 &" DE 12" + vbNewLine & _
		   "NIVEL 01"), vbInformation + vbOKOnly,"BOA!"
		   acertos = acertos+1
		   call carregar_teste	
	elseif tentativa = O then call carregar_ouvir
	else
	cont = cont-1
	msgbox("VOCÊ ERROU!" + vbNewLine & _
	"QUANTIDADE DE ACERTOS:" & acertos &" DE 12" + vbNewLine & _
	"NIVEL 01"), vbInformation + vbOKOnly,":C"
	call carregar_derrota
 end if
 end sub
 
 sub carregar_nivel2
 randomize(second(time))
 nivel2=int(rnd*5) + 1
 select case nivel2
 case 1:
 nivel2= "APARTAMENTO"
 case 2:
 nivel2= "PRÉDIO"
 case 3:
 nivel2= "AMBIGUIDADE"
 case 4:
 nivel2= "BELISCAR"
 case 5:
 nivel2= "RAIO X"
 end select
 call carregar_tentativa2
 end sub
 
 sub carregar_tentativa2()
 audio.speak (""& nivel2 &"")
 tentativa=inputbox("DIGITE A PALAVRA OUVIDA" + vbNewLine + vbNewLine & _
					"NOME DO JOGADOR(A): " & nome &"" + vbNewLine + vbNewLine & _
					"[O] OUVIR NOVAMENTE A PALAVRA" + vbNewLine & _
					"[P] PULAR A PALAVRA UMA ÚNICA VEZ", "SOLETRANDO")
 if tentativa = nivel2 then
	msgbox("PARABÉNS, VOCÊ ACERTOU A PALAVRA!" + vbNewLine & _
		   "QUANTIDADE DE ACERTOS:" & acertos + 1 &" DE 12" + vbNewLine & _
		   "NIVEL 02"), vbInformation + vbOKOnly,"BOA!"
		   acertos = acertos+1
		   call carregar_teste
	else
	cont = cont-1
	msgbox("VOCÊ ERROU!" + vbNewLine & _
	"QUANTIDADE DE ACERTOS:" & acertos &" DE 12" + vbNewLine & _
	"NIVEL 02"), vbInformation + vbOKOnly,":C"
	call carregar_derrota
 end if	   
 end sub
 
 sub carregar_nivel3
 randomize(second(time))
 nivel3=int(rnd*5) + 1
 select case nivel3
 case 1:
 nivel3= "LINGUIÇA"
 case 2:
 nivel3= "TRANQUILO"
 case 3:
 nivel3= "COMBUSTÍVEL"
 case 4:
 nivel3= "AMANHÃ"
 case 5:
 nivel3= "NÚMEROS"
 end select
 call carregar_tentativa3
 end sub
 
 sub carregar_tentativa3()
 audio.speak (""& nivel3 &"")
 tentativa=inputbox("DIGITE A PALAVRA OUVIDA" + vbNewLine + vbNewLine & _
					"NOME DO JOGADOR(A): " & nome &"" + vbNewLine + vbNewLine & _
					"[O] OUVIR NOVAMENTE A PALAVRA" + vbNewLine & _
					"[P] PULAR A PALAVRA UMA ÚNICA VEZ", "SOLETRANDO")
 if tentativa = nivel3 then
	msgbox("PARABÉNS, VOCÊ ACERTOU A PALAVRA!" + vbNewLine & _
		   "QUANTIDADE DE ACERTOS:" & acertos + 1 &" DE 12" + vbNewLine & _
		   "NIVEL 03"), vbInformation + vbOKOnly,"BOA!"
		   acertos = acertos+1
		   call carregar_teste
	else
	cont = cont-1
	msgbox("VOCÊ ERROU!" + vbNewLine & _
	"QUANTIDADE DE ACERTOS:" & acertos &" DE 12" + vbNewLine & _
	"NIVEL 03"), vbInformation + vbOKOnly,":C"
	call carregar_derrota
 end if	   
 end sub
 
 sub carregar_nivel4
 randomize(second(time))
 nivel4=int(rnd*5) + 1
 select case nivel4
 case 1:
 nivel4= "PARALELEPÍPEDO"
 case 2:
 nivel4= "ABDÔMEN"
 case 3:
 nivel4= "HIERÓGLIFO"
 case 4:
 nivel4= "MERCÚRIO"
 case 5:
 nivel4= "MINISSAIA"
 end select
 call carregar_tentativa4
 end sub
 
 sub carregar_tentativa4()
 audio.speak (""& nivel4 &"")
 tentativa=inputbox("DIGITE A PALAVRA OUVIDA" + vbNewLine + vbNewLine & _
					"NOME DO JOGADOR(A): " & nome &"" + vbNewLine + vbNewLine & _
					"[O] OUVIR NOVAMENTE A PALAVRA" + vbNewLine & _
					"[P] PULAR A PALAVRA UMA ÚNICA VEZ", "SOLETRANDO")
 if tentativa = nivel4 then
	msgbox("PARABÉNS, VOCÊ ACERTOU A PALAVRA!" + vbNewLine & _
		   "QUANTIDADE DE ACERTOS:" & acertos + 1 &" DE 12" + vbNewLine & _
		   "NIVEL 04"), vbInformation + vbOKOnly,"BOA!"
		   acertos = acertos+1
		   call carregar_teste
	else
	cont = cont-1
	msgbox("VOCÊ ERROU!" + vbNewLine & _
	"QUANTIDADE DE ACERTOS:" & acertos &" DE 12" + vbNewLine & _
	"NIVEL 04"), vbInformation + vbOKOnly,":C"
	call carregar_derrota
 end if	   
 end sub
 
 sub carregar_teste()
 if acertos < 3 then call carregar_nivel1 end if
 if acertos < 6 then call carregar_nivel2 end if
 if acertos < 9 then call carregar_nivel3 end if
 if acertos < 12 then call carregar_nivel4 end if
 if acertos = 12 then call carregar_vitoria end if
 end sub
 
 'sub carregar_ouvir()
 'audio.speak (""& nivel1 &"")
 'end sub
 
 sub carregar_vitoria()
 resp=msgbox ("PARABÉNS, VOCÊ VENCEU O JOGO!", vbInformation + vbOKOnly, "PARABÉNS") 
 call carregar_derrota
 end sub
 
 sub carregar_derrota()
 resp=msgbox("DESEJA JOGAR NOVAMENTE?",vbquestion + vbyesno,"ATENÇÃO")
 if resp=vbno then
 wscript.quit
 else
 call carregar_audio
 end if	
 end sub
 
 