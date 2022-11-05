dim numero
dim ante,suce
dim resp, vozAtivado, audio

resp=msgbox("Deseja ativar o recurso de voz?", vbquestion+vbyesno,"ATENCAO")
if resp=vbyes then
    vozAtivado = 1
    call carregar_voz
    sub carregar_voz()
    set audio=createobject("SAPI.SPVOICE")
    audio.rate=1
    audio.volume=100
    call numeros
    end sub
else
    vozAtivado = 0
    call numeros
end if

sub numeros()
    numero= cint (inputbox("Digite um numero","AVISO"))
    ante = numero-1
    suce = numero+1
    
    if vozAtivado = 1 then
            audio.speak("Antecessor: " & ante & "" + vbnewline &_ 
        "Sucessor: " & suce & "")
    else
        msgbox("Antecessor: " & ante & "" + vbnewline &_ 
        "Sucessor: " & suce & ""),vbquestion+vbokonly,"Numero"
    end if
    call novoTeste
end sub

sub novoTeste()
resp=msgbox("Deseja realizar novo teste?", vbquestion+vbyesno,"ATENCAO")
if resp=vbyes then
    call numeros
    else
    wscript.quit
end if
end sub