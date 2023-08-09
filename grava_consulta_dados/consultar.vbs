dim db,rs,sql,cpf,nome,resp,executar

call conecta_banco

sub conecta_banco()
set executar=createobject("wscript.shell")
'String de Conexão SQL-Server
set db=createobject("ADODB.Connection") 'Padrão servirá para qualquer banco de dados
db.open ("Provider=SQLOLEDB;Data Source=FL_CRACKEADO;Initial Catalog=ADSVA2_SI;trusted_connection=yes;")
msgbox("Conexão bem sucedida!!!"),vbinformation + vbokonly,"AVISO"
call consultar_dados
end sub

sub consultar_dados()
cpf=clng(inputbox("Digite o CPF do Cliente"))
sql="select * from tb_cadastro where cpf="& cpf &""
set rs=db.execute(sql)
if rs.eof=false then
   resp=msgbox("CPF do Cliente: "& rs.fields(0).value &"" + vbnewline & _
                         "Nome do Cliente: "& rs.fields(1).value &"" + vbnewline & _
						 "Deseja realizar nova Consulta?",vbquestion+vbyesno,"CONSULTA INDIVIDUALIZADA CLIENTES")
   if resp=vbyes then
      call consultar_dados
   else
       executar.run "D:\AULAS2\SI\aula13\menu.vbs"
	   wscript.quit
    end if
else
    msgbox("CPF não existe no cadastro"),vbinformation+vbokonly,"ATENÇÃO"
	call consultar_dados
end if
end sub
   
