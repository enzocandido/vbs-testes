dim n1,n2,n3,media,situacao 'Declaraçao de variaveis
dim resp, audio

call carregar_voz
sub carregar_voz()
set audio=createobject("SAPI.SPVOICE")
audio.rate=1 'Velocidade da fala
audio.volume=100
call entrada_notas 'chamada de funcao
end sub

sub entrada_notas () 
n1= cdbl (inputbox("Digite a nota 1","AVISO"))
n1= cdbl (inputbox("Digite a nota 2","AVISO"))
n1= cdbl (inputbox("Digite a nota 3","AVISO"))

media=round((n1+n2+n3)/3,1)     'Round= Ajusta casas decimais

if media < 4 then
    situacao = "Reprovado"
elseif media >=4 and media <6 then
    situacao = "Recuperaçao"
else 
    situacao = "Aprovado"
end if

'Saida de dados por voz
audio.speak("Media do aluno " & media & ""+ vbnewline &_ 
        "Situacao do aluno" & situacao &"")

'Saida de dados por mensagem
msgbox("Media do Aluno: " & media & "" + vbnewline &_ 
       "Situacao do Aluno: " & situacao & ""),vbquestion+vbokonly,"Rendimento do Aluno"
call novo_calculo
end sub

sub novo_calculo()
resp=msgbox("Deseja realizar novo calculo?", vbquestion+vbyesno,"ATENCAO")
IF resp=vbyes then
    call entrada_notas
    else
    wscript.quit
end if
end sub