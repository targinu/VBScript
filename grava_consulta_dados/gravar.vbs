dim db,rs,sql,cpf,nome,resp,executar

call conecta_banco

sub conecta_banco()
set executar=createobject("wscript.shell")
'String de Conex�o SQL-Server
set db=createobject("ADODB.Connection") 'Padr�o servir� para qualquer banco de dados
'String de Conex�o com o Banco de Dados SQL-SERVER
db.open ("Provider=SQLOLEDB;Data Source=FL_CRACKEADO;Initial Catalog=ADSVA2_SI;trusted_connection=yes;")

'Conex�o com o Banco de Dados MSACCESS
'db.open ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\AULAS2\SI\aula13\ADSVA2_SI.accdb")

msgbox("Conex�o bem sucedida!!!"),vbinformation + vbokonly,"AVISO"
call gravar_dados
end sub

sub gravar_dados()
cpf=clng(inputbox("Digite o CPF do Cliente"))
nome=inputbox("Digite o Nome do Cliente")
sql="select * from tb_cadastro where cpf="& cpf &""
set rs=db.execute(sql)
if rs.eof=false then
   msgbox("CPF: "& cpf &" j� cadastrado!"),vbexclamation + vbokonly,"ATEN��O"
   call gravar_dados
else
   sql="insert into tb_cadastro values ("& cpf &",'"& nome &"')"
   set rs=db.execute(ucase(sql))
   resp=msgbox("Registro Cadastrado com Sucesso!!!" + vbnewline & _
                         "Deseja cadastrar novo Registro?", vbquestion + vbyesno,"ATEN��O")
   if resp=vbyes then
      call gravar_dados
   else
       executar.run "D:\AULAS2\SI\aula13\menu.vbs"
	   wscript.quit
    end if
end if
end sub
   
