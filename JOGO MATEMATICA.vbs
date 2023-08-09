dim num1(10),mat,num2(10), sort1, sort2, sortm, n, conta, qtdacerto, resp

call carregar_nums1
sub carregar_nums1()
num1(1)= "1"
num1(2)= "2"
num1(3)= "3"
num1(4)= "4"
num1(5)= "5"
num1(6)= "6"
num1(7)= "7"
num1(8)= "8"
num1(9)= "9"
num1(10)= "10"
for n=1 to 1 step 1
randomize(second(time))
sort1=int(rnd*10) + 1
next
call carregar_mat
end sub

sub carregar_mat()
sortm=int(rnd*3) + 1
select case sortm
case 1:
sortm= "+"
case 2:
sortm= "-"
case 3:
sortm= "*"
end select
call carregar_nums2
end sub

sub carregar_nums2()
num2(1)= "1"
num2(2)= "2"
num2(3)= "3"
num2(4)= "4"
num2(5)= "5"
num2(6)= "6"
num2(7)= "7"
num2(8)= "8"
num2(9)= "9"
num2(10)= "10"
for n=1 to 1 step 1
randomize(second(time))
sort2=int(rnd*10) + 1
next
call carregar_conta
end sub

sub carregar_conta()
conta=cdbl(inputbox("FAÇA O CALCULO:" + vbNewLine & sort1 &" " & sortm &" " & sort2 &"", "JOGO DA MATEMATICA"))

if sortm = "+" then mat = 1
if sortm = "-" then mat = 2
if sortm = "*" then mat = 3

	if mat = 1 and conta = sort1 + sort2 then qtdacerto = qtdacerto +1
	if mat = 2 and conta = sort1 - sort2 then qtdacerto = qtdacerto +1
	if mat = 3 and conta = sort1 * sort2 then qtdacerto = qtdacerto +1
	
	if mat = 1 and conta = sort1 + sort2 then call carregar_acerto
	if mat = 2 and conta = sort1 - sort2 then call carregar_acerto
	if mat = 3 and conta = sort1 * sort2 then call carregar_acerto
	
	if mat = 1 and conta <> sort1 + sort2 then msgbox("VOCÊ ERROU!" +vbNewLine &"QUANTIDADE DE ACERTOS: " & qtdacerto &"") ,vbInformation + vbOKOnly , ":(" 
	if mat = 2 and conta <> sort1 - sort2 then msgbox("VOCÊ ERROU!" +vbNewLine &"QUANTIDADE DE ACERTOS: " & qtdacerto &"") ,vbInformation + vbOKOnly , ":(" 
	if mat = 3 and conta <> sort1 * sort2 then msgbox("VOCÊ ERROU!" +vbNewLine &"QUANTIDADE DE ACERTOS: " & qtdacerto &"") ,vbInformation + vbOKOnly , ":(" 
	
resp=msgbox("Deseja encerrar?",vbquestion + vbyesno,"ATENÇÃO")
if resp=vbyes then
wscript.quit
else
call carregar_nums1
end if	
	 
end sub

sub carregar_acerto()
msgbox("VOCÊ ACERTOU!" +vbNewLine &"QUANTIDADE DE ACERTOS: " & qtdacerto &"") ,vbInformation + vbOKOnly, "PARABÉNS" 
carregar_nums1
end sub