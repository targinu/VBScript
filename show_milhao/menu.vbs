dim op,executar,resp
call carregar_menu
sub carregar_menu()
set executar=createobject("wscript.shell")
op=inputbox("[C]CADASTRAR PERGUNTA" + vbnewline & _
                    "[J]JOGAR AGORA" + vbnewline & _
					"[F]FINALIZAR PROGRAMA","CONHECIMENTOS GERAIS - MENU PRINCIPAL")
select case op
                   case "C","c":
				            executar.run "D:\AULAS2\SI\show_milhao\cadastrar.vbs"
				            wscript.quit
				   case "J","j":
				            executar.run "D:\AULAS2\SI\show_milhao\questoes.vbs"
				            wscript.quit
				   case "F","f":
				            resp=msgbox("Deseja Encerrar?",vbquestion+vbyesno,"ATENÇÃO")
							if resp=vbyes then
							   wscript.quit
							end if
				    case else
					        msgbox("Opção Inválida!"),vbexclamation+vbokonly,"ATENÇÃO"
							call carregar_menu
end select
end sub