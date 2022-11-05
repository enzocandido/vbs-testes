dim numeros(30)

numeros(0)=0
numeros(1)=1

numero=inputbox("Digite a quantidade de series numerica:")

for laco = 2 to numero
    numeros(laco)=numeros(laco-1)+numeros(laco-2)
next

for laco = 1 to numero
    msgbox numeros(laco)
next