import { Bot, InlineKeyboard, webhookCallback } from "grammy";
import { chunk } from "lodash";
import express from "express";
import { applyTextEffect, Variant } from "./textEffects";

import type { Variant as TextEffectVariant } from "./textEffects";

// Create a bot using the Telegram token
const bot = new Bot(process.env.TELEGRAM_TOKEN || "");

/********************** Cuando mandás /Start **********************/

bot.start((ctx) => {

  try {
    ctx.reply('Bienvenido al sistema de guardias de ACN\nPor favor, presione /help para desplegar las funciones disponibles. ');
    console.log(ctx.chat.id);
  } catch (error) {
    console.error(error);
  }
console.log(ctx.chat.id);
});



/********************** Cuando mandás /Help **********************/

bot.help((ctx) => {
  try {
    ctx.reply('Enviá /guardia para conocer quién se encuentra de guardia actualmente.\n' +
              'Enviá /proxguardia para conocer quién se encontrará de guardia la semana próxima.\n' +
              'Enviá /asociarCuenta seguido de su número de rotación para relacionar tu cuenta de Telegram y recibir un mensaje cuando estés de Guardia.\n' +
              'Enviá /eliminarCuenta para eliminar la relación de tu cuenta de Telegram.');
  } catch (error) {
    console.error(error);
  }
});

/********************** Lectura del analista **********************/

bot.command('guardia', (ctx) => {
  try {
    ctx.reply(leerAnalista());
  } catch (error) {
    console.error(error);
  }
});


/***************** Lectura del próximo analista ********************/


bot.command('proxguardia', (ctx) => {
  try {
    ctx.reply(leerProxAnalista());
  } catch (error) {
    console.error(error);
  }
});

/********************** Para asociar la cuenta **********************/
bot.command('asociarCuenta', (ctx) => {
  try {
    const respuesta = cargarUsuario(ctx);
    ctx.reply(respuesta);
  } catch (error) {
    console.error(error);
  }
});

/********************** Para eliminar la cuenta **********************/
bot.command('eliminarCuenta', (ctx) => {
  try {
    const respuesta = eliminarUsuario(ctx);
    ctx.reply(respuesta);
  } catch (error) {
    console.error(error);
  }
});

/********************** Función que lee el calendario **********************/

function leerExcel (ruta){
    const workbook = Excel.readFile(ruta);
    const workbookSheets = workbook.SheetNames;
    const sheet = workbookSheets[0];
    const dataExcel = Excel.utils.sheet_to_json(workbook.Sheets[sheet]);
    
    return dataExcel;
    
//    console.log(dataExcel);
//    console.log(dataExcel[1]['Analista']);

}



//console.log('Hoy está de guardia '  + hoja[1]['Analista']);

// Convertirlo a una fecha JavaScript
//const date = new Date((hoja[1]['Start Day'] - 1) * 24 * 60 * 60 * 1000 + Date.UTC(1900, 0, 1));


const hoja = leerExcel("baseDatos.xlsx");



/********************** Función que lee el analista actual **********************/

function leerAnalista(){

let date = new Date();
date = date.toISOString().slice(0, 10);

console.log(date);

  for (const item of hoja) {
    // Convierte las fechas en el objeto actual a objetos Date
    let startDate = new Date((item['Start Day']- 1) * 24 * 60 * 60 * 1000 + Date.UTC(1900, 0, 1));
    let endDate = new Date((item['End Day']- 1) * 24 * 60 * 60 * 1000 + Date.UTC(1900, 0, 1));
    startDate = startDate.toISOString().slice(0, 10);
    endDate = endDate.toISOString().slice(0, 10);

    // Comprueba si 'date' está dentro del rango de fechas del objeto actual
    if (date >= startDate && date <= endDate) {
      return ('Está de guardia ' + item['Analista'] + ' hasta el día ' + endDate );
    } 
  }
}

/********************** Función que lee el próximo analista **********************/

function leerProxAnalista(){

  let date = new Date();
  date.setDate(date.getDate() + 7);
  date = date.toISOString().slice(0, 10);
  
  console.log(date);
  
    for (const item of hoja) {
      // Convierte las fechas en el objeto actual a objetos Date
      let startDate = new Date((item['Start Day']- 1) * 24 * 60 * 60 * 1000 + Date.UTC(1900, 0, 1));
      let endDate = new Date((item['End Day']- 1) * 24 * 60 * 60 * 1000 + Date.UTC(1900, 0, 1));
      startDate = startDate.toISOString().slice(0, 10);
      endDate = endDate.toISOString().slice(0, 10);
  
      // Comprueba si 'date' está dentro del rango de fechas del objeto actual
      if (date >= startDate && date <= endDate) {
        return ('La semana que viene está de guardia ' + item['Analista'] + ' hasta el día ' + endDate );
      } 
    }
  }

/*
let date = new Date();

console.log (date);
date.setDate(date.getDate() + 7)
console.log (date);
console.log (date.getHours());
console.log (date.getMinutes());
*/



/********************** Evento que lee el analista nuevo **********************/

function verificarYEnviarMensaje() {
  const ahora = new Date();
  
  // Verificar si hoy es viernes (5 es el código para viernes)
  if (ahora.getDay() === 5 && ahora.getHours() === 18) {
    bot.telegram.sendMessage('1656162437', 'A partir de este momento, te toca la guardia! Que tengas buen fin de semana.');
    console.log("¡Es viernes a las 18:00! Tu mensaje aquí.");
  }
}

// Configurar el intervalo para verificar cada minuto
setInterval(verificarYEnviarMensaje, 60 * 1000); // Cada 1 minuto


const now = new Date();
console.log(now.getHours());


/********************** Función que asocia analista **********************/

function cargarUsuario(ctx){

const XLSX = require('xlsx');
const fs = require('fs');

// Ruta al archivo Excel existente
const filePath = 'baseDatos.xlsx';

// Leer el archivo Excel
const workbook = XLSX.readFile(filePath);

// Obtener la hoja específica (por nombre)
const worksheet = workbook.Sheets['Usuario'];

//console.log('WORKSHEET: '+ worksheet);

// Crear un objeto con los datos que deseas agregar
/*const datosAgregados = [
  { id: ctx.chat.id, rotacion: ctx.message.text.slice(6,7) }
];
*/

const data = XLSX.utils.sheet_to_json(worksheet);



console.log(ctx.chat.id);
console.log(ctx.message.text);


const id = ctx.chat.id;
const rotacion = ctx.message.text;

console.log(id);
console.log(rotacion.slice(15,16));

if (rotacion.slice(15,16) == '1' || rotacion.slice(15,16) == '2' || rotacion.slice(15,16) == '3' || rotacion.slice(15,16) == '4'){

let idExistente = false;
for (const item of data) {
  if (item.id === id) {
    idExistente = true;
    break;
  }
}

for (const item of data) {
  if (item.rotacion === rotacion) {
    idExistente = true;
    break;
  }
}

if (!idExistente) {
  // Agregar el nuevo dato solo si el id no existe
data.push({ id: id, rotacion: rotacion.slice(15,16) });


console.log ('DATOS AGREGADOS: ' + data);


// Crear una nueva hoja de Excel con los datos actualizados
const nuevaWorksheet = XLSX.utils.json_to_sheet(data);

console.log(nuevaWorksheet);

// Reemplazar la hoja original con la nueva hoja
workbook.Sheets['Usuario'] = nuevaWorksheet;

console.log(workbook.Sheets['Usuario']);

/*
// Crear un buffer de datos a partir del libro modificado
const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });

// Escribir el buffer de datos en el archivo existente
fs.writeFileSync(filePath, excelBuffer);
*/



// Guardar el libro con los datos actualizados
XLSX.writeFile(workbook, filePath);

return('Se asoció el número de rotación a esta cuenta.');
/*
console.log('Datos agregados a la hoja en el archivo Excel existente.');

*/
}
else{
  return ('Ya existe ese número de rotación o la cuenta ya está asociada.');
}
}
else{
  return ('Ingrese un número de rotación válido.\nEjemplo /asociarCuenta 2');
}
}




/********************** Función que elimina analista **********************/


function eliminarUsuario(ctx) {
  const XLSX = require('xlsx');
  const fs = require('fs');

  // Ruta al archivo Excel existente
  const filePath = 'baseDatos.xlsx';

  // Leer el archivo Excel
  const workbook = XLSX.readFile(filePath);

  // Obtener la hoja específica (por nombre)
  const sheetName = 'Usuario'; // Reemplaza 'Usuario' con el nombre de tu hoja
  const worksheet = workbook.Sheets[sheetName];

  const data = XLSX.utils.sheet_to_json(worksheet);

  console.log(ctx.chat.id);

  const id = ctx.chat.id;

  // Buscar el índice del registro que deseas eliminar
  let indiceAEliminar = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i].id === id) {
      indiceAEliminar = i;
      break;
    }
  }

  if (indiceAEliminar !== -1) {
    // Eliminar el registro basado en el índice encontrado
    data.splice(indiceAEliminar, 1);

    // Crear una nueva hoja de Excel con los datos actualizados
    const nuevaWorksheet = XLSX.utils.json_to_sheet(data);

    // Reemplazar la hoja original con la nueva hoja
    workbook.Sheets[sheetName] = nuevaWorksheet;

    // Guardar el libro con los datos actualizados
    XLSX.writeFile(workbook, filePath);

    return 'Registro eliminado con éxito.';
  } else {
    return 'No se encontró ningún registro con ese ID.';
  }
}


/********************** Función que busca y envia al analista **********************/

// Función para buscar y enviar mensajes a los analistas
function buscarYEnviarAnalista() {

  const XLSX = require('xlsx');
  const fs = require('fs');

  // Ruta al archivo Excel existente
  const filePath = 'baseDatos.xlsx';

  // Leer el archivo Excel
  const workbook = XLSX.readFile(filePath);
  // Obtener la fecha actual
  const ahora = new Date();

  // Verificar si hoy es viernes (5 es el código para viernes) y la hora es 18:00
  if (ahora.getDay() === 5 && ahora.getHours() === 18) {
    // Leer el archivo Excel para obtener la información de los analistas
     const usuario = workbook.Sheets['Usuario'];
    const calendario = workbook.Sheets['2023'];

    // Obtener todos los datos de la hoja
    const dataUsuario = XLSX.utils.sheet_to_json(usuario);
    const dataCalendario = XLSX.utils.sheet_to_json(calendario);

    //Buscar el analista
    let date = new Date();
    let rotacion = null;
    date = date.toISOString().slice(0, 10);

      for (const item of dataCalendario) {
        // Convierte las fechas en el objeto actual a objetos Date
        let startDate = new Date((item['Start Day']- 1) * 24 * 60 * 60 * 1000 + Date.UTC(1900, 0, 1));
        let endDate = new Date((item['End Day']- 1) * 24 * 60 * 60 * 1000 + Date.UTC(1900, 0, 1));
        startDate = startDate.toISOString().slice(0, 10);
        endDate = endDate.toISOString().slice(0, 10);
    
        // Comprueba si 'date' está dentro del rango de fechas del objeto actual
        if (date >= startDate && date <= endDate) {
          rotacion = item['Rotacion']
        } 
      }

      console.log(rotacion);


    // Iterar a través de los datos para buscar al analista de guardia
    let analistaDeGuardia = null;
    for (const item of dataUsuario) {
      // Aquí puedes realizar la lógica para buscar al analista de guardia
      // Puedes comparar la fecha actual con las fechas en los datos, buscar por ID, etc.
      // Cuando encuentres al analista de guardia, asigna su información a 'analistaDeGuardia'
     
      console.log(item);
      console.log(rotacion);
      // Ejemplo: Si el ID del analista coincide con el del chat actual
      if (item.rotacion === rotacion) {
        analistaDeGuardia = item;
        break;
      }
    }
    console.log(analistaDeGuardia);
    // Si se encontró al analista de guardia, envía el mensaje
    if (analistaDeGuardia) {
      const mensaje = `A partir de este momento, te toca la guardia! Que tengas buen fin de semana.`;
      const chatId = analistaDeGuardia.id;
      bot.telegram.sendMessage(chatId, mensaje);
      console.log(`Mensaje enviado al analista ${analistaDeGuardia.nombre}`);
    }
  }
}

// Configurar el intervalo para verificar cada minuto
setInterval(buscarYEnviarAnalista, 3600 * 1000); // Cada 1 hora

/*
// Handle the /yo command to greet the user
bot.command("yo", (ctx) => ctx.reply(`Yo ${ctx.from?.username}`));

// Handle the /effect command to apply text effects using an inline keyboard
type Effect = { code: TextEffectVariant; label: string };
const allEffects: Effect[] = [
  {
    code: "w",
    label: "Monospace",
  },
  {
    code: "b",
    label: "Bold",
  },
  {
    code: "i",
    label: "Italic",
  },
  {
    code: "d",
    label: "Doublestruck",
  },
  {
    code: "o",
    label: "Circled",
  },
  {
    code: "q",
    label: "Squared",
  },
];

const effectCallbackCodeAccessor = (effectCode: TextEffectVariant) =>
  `effect-${effectCode}`;

const effectsKeyboardAccessor = (effectCodes: string[]) => {
  const effectsAccessor = (effectCodes: string[]) =>
    effectCodes.map((code) =>
      allEffects.find((effect) => effect.code === code)
    );
  const effects = effectsAccessor(effectCodes);

  const keyboard = new InlineKeyboard();
  const chunkedEffects = chunk(effects, 3);
  for (const effectsChunk of chunkedEffects) {
    for (const effect of effectsChunk) {
      effect &&
        keyboard.text(effect.label, effectCallbackCodeAccessor(effect.code));
    }
    keyboard.row();
  }

  return keyboard;
};

const textEffectResponseAccessor = (
  originalText: string,
  modifiedText?: string
) =>
  `Original: ${originalText}` +
  (modifiedText ? `\nModified: ${modifiedText}` : "");

const parseTextEffectResponse = (
  response: string
): {
  originalText: string;
  modifiedText?: string;
} => {
  const originalText = (response.match(/Original: (.*)/) as any)[1];
  const modifiedTextMatch = response.match(/Modified: (.*)/);

  let modifiedText;
  if (modifiedTextMatch) modifiedText = modifiedTextMatch[1];

  if (!modifiedTextMatch) return { originalText };
  else return { originalText, modifiedText };
};

bot.command("effect", (ctx) =>
  ctx.reply(textEffectResponseAccessor(ctx.match), {
    reply_markup: effectsKeyboardAccessor(
      allEffects.map((effect) => effect.code)
    ),
  })
);

// Handle inline queries
const queryRegEx = /effect (monospace|bold|italic) (.*)/;
bot.inlineQuery(queryRegEx, async (ctx) => {
  const fullQuery = ctx.inlineQuery.query;
  const fullQueryMatch = fullQuery.match(queryRegEx);
  if (!fullQueryMatch) return;

  const effectLabel = fullQueryMatch[1];
  const originalText = fullQueryMatch[2];

  const effectCode = allEffects.find(
    (effect) => effect.label.toLowerCase() === effectLabel.toLowerCase()
  )?.code;
  const modifiedText = applyTextEffect(originalText, effectCode as Variant);

  await ctx.answerInlineQuery(
    [
      {
        type: "article",
        id: "text-effect",
        title: "Text Effects",
        input_message_content: {
          message_text: `Original: ${originalText}
Modified: ${modifiedText}`,
          parse_mode: "HTML",
        },
        reply_markup: new InlineKeyboard().switchInline("Share", fullQuery),
        url: "http://t.me/EludaDevSmarterBot",
        description: "Create stylish Unicode text, all within Telegram.",
      },
    ],
    { cache_time: 30 * 24 * 3600 } // one month in seconds
  );
});

// Return empty result list for other queries.
bot.on("inline_query", (ctx) => ctx.answerInlineQuery([]));

// Handle text effects from the effect keyboard
for (const effect of allEffects) {
  const allEffectCodes = allEffects.map((effect) => effect.code);

  bot.callbackQuery(effectCallbackCodeAccessor(effect.code), async (ctx) => {
    const { originalText } = parseTextEffectResponse(ctx.msg?.text || "");
    const modifiedText = applyTextEffect(originalText, effect.code);

    await ctx.editMessageText(
      textEffectResponseAccessor(originalText, modifiedText),
      {
        reply_markup: effectsKeyboardAccessor(
          allEffectCodes.filter((code) => code !== effect.code)
        ),
      }
    );
  });
}

// Handle the /about command
const aboutUrlKeyboard = new InlineKeyboard().url(
  "Host your own bot for free.",
  "https://cyclic.sh/"
);

// Suggest commands in the menu
bot.api.setMyCommands([
  { command: "yo", description: "Be greeted by the bot" },
  {
    command: "effect",
    description: "Apply text effects on the text. (usage: /effect [text])",
  },
]);

// Handle all other messages and the /start command
const introductionMessage = `Hello! I'm a Telegram bot.
I'm powered by Cyclic, the next-generation serverless computing platform.

<b>Commands</b>
/yo - Be greeted by me
/effect [text] - Show a keyboard to apply text effects to [text]`;

const replyWithIntro = (ctx: any) =>
  ctx.reply(introductionMessage, {
    reply_markup: aboutUrlKeyboard,
    parse_mode: "HTML",
  });

bot.command("start", replyWithIntro);
bot.on("message", replyWithIntro);
*/

// Start the server
if (process.env.NODE_ENV === "production") {
  // Use Webhooks for the production server
  const app = express();
  app.use(express.json());
  app.use(webhookCallback(bot, "express"));

  const PORT = process.env.PORT || 3000;
  app.listen(PORT, () => {
    console.log(`Bot listening on port ${PORT}`);
  });
} else {
  // Use Long Polling for development
  bot.start();
}
