let chassis = [];
let encontrados = new Set();

const total = document.getElementById("total");
const ok = document.getElementById("ok");
const faltam = document.getElementById("faltam");
const barra = document.getElementById("barra");
const mensagem = document.getElementById("mensagem");
const lista = document.getElementById("lista");

document.getElementById("arquivo").addEventListener("change", carregarExcel);

function carregarExcel(e){

    const reader = new FileReader();

    reader.onload = function(evt){

        const data = new Uint8Array(evt.target.result);

        const workbook = XLSX.read(data,{type:"array"});

        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        const json = XLSX.utils.sheet_to_json(sheet);

        chassis = json.map(item=>({

            chassi:String(item["Chassi"]).trim(),

            modelo:item["Mod /Versão"] || "",

            data:item["Data da Fatura"] || "",

            status:item["Status Tropicalização"] || "",

            madt:item["Data MADT"] || "",

            encontrado:false

        }));

        atualizarPainel();

        mensagem.innerHTML="✅ Planilha carregada.";
    };

    reader.readAsArrayBuffer(e.target.files[0]);

}

function atualizarPainel(){

    total.innerHTML = chassis.length;

    ok.innerHTML = encontrados.size;

    faltam.innerHTML = chassis.length - encontrados.size;

    barra.max = chassis.length;

    barra.value = encontrados.size;

}

function conferirChassi(codigo){

    codigo = codigo.trim();

    const veiculo = chassis.find(v=>v.chassi===codigo);

    if(!veiculo){

        mensagem.innerHTML="🔴 Veículo não encontrado.";

        mensagem.style.color="red";

        return;

    }

    if(veiculo.encontrado){

        mensagem.innerHTML="🟡 Veículo já conferido.";

        mensagem.style.color="orange";

        return;

    }

    veiculo.encontrado=true;

    encontrados.add(codigo);

    atualizarPainel();

    mensagem.style.color="green";

    mensagem.innerHTML=
`🟢 VEÍCULO ENCONTRADO<br><br>
<b>${veiculo.modelo}</b><br><br>
${veiculo.chassi}`;

}

document.getElementById("faltantesBtn").onclick=function(){

    lista.innerHTML="";

    chassis
    .filter(v=>!v.encontrado)
    .forEach(v=>{

        const li=document.createElement("li");

        li.innerHTML=
        `<b>${v.chassi}</b><br>${v.modelo}`;

        lista.appendChild(li);

    });

};
