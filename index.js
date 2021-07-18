const ExcelJS = require('exceljs');
const cheerio = require('cheerio');
const fetch = require('node-fetch');
const fs = require('fs');

const workbook = new ExcelJS.Workbook();

// Returns a Promise that resolves after "ms" Milliseconds
const timer = (ms) => new Promise((res) => setTimeout(res, ms));

const arrayToString = (arr) => {
  let defsString = '';

  arr.forEach((defObj) => {
    const def = defObj.i;
    const role = def.slice(0, def.split('').indexOf('.') + 1);
    defsString += `${def.replace(role, `<i>${role}</i>`)}<br />`;
  });
  return defsString;
};

const readExcelList = async () => {
  const list = [];

  try {
    await workbook.xlsx.readFile('./list.xlsx');
    const ws = workbook.getWorksheet('Sheet1');
    const column = ws.getColumn('A');
    column.eachCell((cell, rowNumber) => {
      const word = cell.value;
      list.push(word);
    });
  } catch (error) {
    console.log(error);
    throw error;
  }

  return list;
};

const getIdDefs = async (q, id) => {
  let idDefs;

  try {
    await timer(2000 * id);
    // Bahasa indonesia
    const uDictionaryURL = `https://inter.youdao.com/intersearch?q=${q}&from=en&to=id&interversion=23`;
    const uDictionaryResponse = await fetch(uDictionaryURL);
    const uDictionaryJSON = await uDictionaryResponse.json();
    const idDefsArray = uDictionaryJSON.data.eh.trs;
    idDefs = arrayToString(idDefsArray);
  } catch (error) {
    idDefs = '';
    console.error(error);
    throw error;
  }

  return idDefs;
};

const getSentenceExamples = async (q, id) => {
  // Word in sentences
  try {
    await timer(2000 * id);
    const exampleURL = `https://corpus.vocabulary.com/api/1.0/examples.json?maxResults=10&query=${q}&startOffset=0&domain=F`;
    const exampleResponse = await fetch(exampleURL);
    const json = await exampleResponse.json();
    const sentences = json.result.sentences;
    let example = '';
    sentences.forEach(({ sentence, offsets }) => {
      const wordUsed = sentence.slice(offsets[0], offsets[1]);
      const formatSentence = sentence.replace(wordUsed, `<i><b>${wordUsed}</b></i>`);
      example = `${example + formatSentence}<br /><br />`;
    });

    return example;
  } catch (error) {
    console.error(error);
    throw error;
  }
};

const getWordDefinitions = async (q, id) => {
  try {
    await timer(2000 * id);

    // Word definition
    const url = `https://www.vocabulary.com/dictionary/${q}`;
    const response = await fetch(url);
    const html = await response.text();

    // Use cheerio
    const $ = await cheerio.load(html, {
      xml: {
        normalizeWhitespace: true,
      },
    });
    const short = await $('.short').html();
    const long = await $('.long').html();
    const definition = await $('.sense > .definition');

    let defs = '';
    await definition.map((i, el) => {
      const fullText = $(el).text();
      const role = $(el).children('div').text();
      defs = `${defs + fullText.replace(role, `<i>${role}</i>`)}<br />`;

      return null;
    });

    return {
      short, long, defs, url,
    };
  } catch (error) {
    console.error(error);
    throw error;
  }
};

const getDetails = async (list) => {
  const details = await Promise.all(list.map(async (word, id) => {
    try {
      const query = word.trim().toLowerCase();

      const idDefs = await getIdDefs(query, id);
      const example = await getSentenceExamples(query, id);
      const {
        short, long, defs, url,
      } = await getWordDefinitions(query, id);

      console.log({
        word,
        idDefs: idDefs || '',
        id,
        short,
        long,
        defs,
        example,
        url: `<a href=${url}>vocabulary.com</a>`,
      });

      if(defs.length === 0) {
        return null;
      }

      return {
        word,
        idDefs: idDefs || '',
        id,
        short,
        long,
        defs,
        example,
        url: `<a href=${url}>vocabulary.com</a>`,
      };
    } catch (error) {
      return null;
    }
  }));

  return details;
};

const generateTSV = async () => {
  const list = await readExcelList();

  // Read word and find definition on vocabulary.com using cheerio
  const details = await getDetails(list);

  let final = '';
  details.forEach((word) => {
    let line = '';

    if(word) {
      // eslint-disable-next-line no-restricted-syntax
      for (const [key, value] of Object.entries(word)) {
        if(key !== 'id' && key !== 'url') {
          line = `${line + value}\t`;
        }else if(key !== 'id' && key === 'url') {
          line = `${line + value}\n`;
        }
      }
    }

    final += line;
  });

  if(final) {
    await fs.writeFile('./export.txt', final, (err) => {
      if(err) throw err;
      console.log("it's saved");
    });
  }else{
    console.log('No result');
  }
};

generateTSV();
