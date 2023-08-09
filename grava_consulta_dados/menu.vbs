dim op,executar,resp
call carregar_menu
sub carregar_menu()
set executar=createobject("wscript.shell")
op=inputbox("[G]ravar Dados" + vbnewline & _
                    "[C]onsultar Dados" + vbnewline & _
					"[F]inalizar Script","ESCOLHA UMA OPÇÃO")
select case op
                   case "G","g":
				            executar.run "D:\AULAS2\SI\aula13\gravar.vbs"
				            wscript.quit
				   case "C","c":
				            executar.run "D:\AULAS2\SI\aula13\consultar.vbs"
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
				 
					