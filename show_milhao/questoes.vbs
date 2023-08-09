dim db,rs,sql,cpf,nome,resp,executar, numero

call conecta_banco

sub conecta_banco()
set executar=createobject("wscript.shell")
set db=createobject("ADODB.Connection")
db.open ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\AULAS2\SI\show_milhao\show_milhao.accdb")
msgbox("Conexão bem sucedida!!!"),vbinformation + vbokonly,"AVISO"
call gerar_pergunta
end sub

sub gerar_pergunta()
randomize(second(time))
n=int(rnd*502)+1
sql="select * from tb_questoes where numero="& numero &""
set rs=db.execute(sql)
if rs.eof=false then
   resp=msgbox("PERGUNTA: "& rs.fields(0).value &"" + vbnewline & _
                         "A) "& rs.fields(1).value &"" + vbnewline & _
						 "B) "& rs.fields(2).value &"" + vbnewline & _
						 "C) "& rs.fields(3).value &"" + vbnewline & _
						 "D) "& rs.fields(4).value &"" + vbnewline & _
						 "Deseja realizar nova Consulta?",vbquestion+vbyesno,"CONSULTA INDIVIDUALIZADA CLIENTES")
   if resp=vbyes then
      call consultar_dados
   else
       executar.run "C:\Users\Humberto\OneDrive\Desktop\menu.vbs"
	   wscript.quit
    end if
else
     msgbox("CPF não existe no cadastro"),vbinformation+vbokonly,"ATENÇÃO"
	 call consultar_dados
end if
end sub
   
