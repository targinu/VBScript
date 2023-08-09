dim numero,

call conecta_banco

sub conecta_banco()
set executar=createobject("wscript.shell")
set db=createobject("ADODB.Connection")
db.open ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\AULAS2\SI\show_milhao\show_milhao.mdb")
msgbox("Conexão bem sucedida!!!"),vbinformation + vbokonly,"AVISO"
call cadastrar_pergunta_pergunta
end sub

sub cadastrar_pergunta