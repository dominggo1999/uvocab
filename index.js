const ExcelJS = require('exceljs');
const cheerio = require('cheerio');
const fetch = require('node-fetch');
const fs = require('fs');

const workbook = new ExcelJS.Workbook();

const readExcelList = async () => {
  const list = [];
  await workbook.xlsx.readFile('./list.xlsx');
  const ws = workbook.getWorksheet('Sheet1');
  const column = ws.getColumn('A');
  column.eachCell((cell, rowNumber) => {
    const word = cell.value;
    list.push(word);
  });

  return list;
};

const getDetails = async (list) => {
  const details = await Promise.all(list.map(async (word, id) => {
    const exampleURL = `https://corpus.vocabulary.com/api/1.0/examples.json?maxResults=10&query=${word}&startOffset=0&domain=F`;
    const exampleResponse = await fetch(exampleURL);
    const json = await exampleResponse.json();
    const sentences = json.result.sentences;
    let example = '';
    sentences.forEach(({ sentence, offsets }) => {
      const wordUsed = sentence.slice(offsets[0], offsets[1]);
      const formatSentence = sentence.replace(wordUsed, `<i><b>${wordUsed}</b></i>`);
      example = `${example + formatSentence}<br /><br />`;
    });

    const url = `https://www.vocabulary.com/dictionary/${word}`;
    const response = await fetch(url);
    const html = await response.text();

    // Use cheerio
    const $ = await cheerio.load(html, {
      xml: {
        normalizeWhitespace: true,
      },
    });
    const long = await $('.long').html();
    const short = await $('.short').html();
    const definition = await $('.sense > .definition');

    let defs = '';
    await definition.map((i, el) => {
      const fullText = $(el).text();
      const role = $(el).children('div').text();
      defs = `${defs + fullText.replace(role, `<i>${role}</i>`)}<br />`;

      return null;
    });

    if(defs.length === 0) {
      return null;
    }

    return {
      word,
      id,
      short,
      long,
      defs,
      example,
      url: `<a href=${url}>vocabulary.com</a>`,
    };
  }));

  return details;
};

const generateExel = async () => {
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

generateExel();
