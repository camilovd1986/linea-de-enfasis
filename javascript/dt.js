function ocultartablas(){
  document.getElementById("divtablas").style.display = "none"
  document.getElementById("divbotonchiller").style.display = "none"
  document.getElementById("divbotonmostrar").style.display = "none"
}

function paso1() {
    // Pedimos al usuario que introduzca un valor
    let factor;
    do {
    factor = prompt("Introduce un valor de factor de servicio entre 1 y 3:");
    } while (factor === "" || factor === null);
    let caudal;
    do{
      caudal = prompt("Introduce un valor de caudal en m3:");
    } while(caudal === "" || caudal === null);
    let tentrada;
    do {
      tentrada = prompt("Introduce un valor de temperatura de entrada:");
    } while (tentrada === "" || tentrada === null);
    let tsalida;
    do {
      tsalida = prompt("Introduce un valor de temperatura de salida:");
    } while (tsalida === "" || tsalida === null);
    
    // Convertimos el valor introducido en un número y realizamos el cálculo
    let resultado = parseFloat(caudal*(tentrada-tsalida)*factor*1000*0.0003069);
    // Mostramos el resultado en una ventana emergente
    alert("El tamaño del distrito termico es: " + resultado);
    alert("para continuar presione boton chillers")
    document.getElementById("divbotonchiller").style.display = "block"

  }
  function paso2(){
    // Pedimos al usuario que introduzca un valor
    let c500;
    do {
      c500 = prompt("Introduce cantidad de chiller centrifugo de 500TR:");
    } while (c500 === "" || c500 === null);
    let c750;
    do {
      c750 = prompt("Introduce cantidad de chiller centrifugo de 750TR:");
    } while (c750 === "" || c750 === null);
    let c1000;
    do {
      c1000 = prompt("Introduce cantidad de chiller centrifugo de 1000TR:");
    } while (c1000 === "" || c1000 === null);
    let a500;
    do {
      a500 = prompt("Introduce cantidad de chiller de absorción de 500TR:");
    } while (a500 === "" || a500 === null);
    let a750;
    do {
      a750 = prompt("Introduce cantidad de chiller de absorción de 750TR:");
    } while (a750 === "" || a750 === null);
    let a1000;
    do {
      a1000 = prompt("Introduce cantidad de chiller de absorción de 1000TR:");
    } while (a1000 === "" || a1000 === null);
    
    // Convertimos el valor introducido en un número y realizamos el cálculo
    let totalc= Number(500*c500)+(750*c750)+(1000*c1000)
    let totala= Number(500*a500)+(750*a750)+(1000*a1000)
    let totales = totala + totalc
    alert("El tamaño del distrito termico es: " + totales);
    //alert("Para continuar presione Imprimir tablas")
    document.getElementById("divbotonmostrar").style.display = "block"
    mostrartabla(totalc,totala)

  }

  function mostrartabla(totalc,totala){

  
    rpc= Number(totalc*0.3190995427365);
    gc= Number(totalc*511.13199046407)/1000;	
    cc= Number(totalc*0.0035174111853)*(1925000/0.88);	
    oc= Number(cc*0.03);	    	
    capexc= Number(totalc*0.0035174111853);	
    ftc= Number(capexc*1000000);	
    ec= Number(capexc*1700000);	
    bc=Number(capexc*2000000);

    // Mostramos los valores en la tabla


    document.getElementById("rpc").innerHTML=rpc;
    document.getElementById("capexc").innerHTML =capexc;
    document.getElementById("gc").innerHTML = gc;
    document.getElementById("cc").innerHTML = cc;
    document.getElementById("oc").innerHTML = oc;
    document.getElementById("ftc").innerHTML = ftc;
    document.getElementById("ec").innerHTML = ec;
    document.getElementById("bc").innerHTML = bc;
    document.getElementById("totalc").innerHTML=totalc;

  //function absorción(totala){
   ga= Number((totala*511.13199046407)/1000);		
   ca= Number((totala * 0.0035174111853)*(1925000/0.88));		
   oa=Number(ca*0.03);	
   capexa= Number(totala*0.0035174111853);		
   fta= Number((capexa*1000000)*1.015);		
   ba= Number(capexa*2000000);

   // Mostramos los valores en la tabla
   document.getElementById("ga").innerHTML = ga;
   document.getElementById("ca").innerHTML = ca;
   document.getElementById("oa").innerHTML = oa;
   document.getElementById("capexa").innerHTML = capexa;
   //document.getElementById("fta").innerHTML = fta;
   document.getElementById("ba").innerHTML = ba;
   document.getElementById("totala").innerHTML=totala;

  
    centri.style.display = "block";
    absor.style.display = "block";
  //}
} 
  //}

