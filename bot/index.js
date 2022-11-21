const puppeteer = require("puppeteer");
const XLSX = require("xlsx");
let workbook = XLSX.readFile("./customers.xlsx");
let worksheet = workbook.Sheets[workbook.SheetNames[0]];

// ** variaveis ** //
const selectOrgMilitar =
  "#q-app > div.container > div.q-layout > div.layout-view > div > div.row.gutter-xs > div.col-md-12.col-xs-12 > div > div.q-if-inner.col.column.q-popup--skip > div > div";
const btnAgendamento =
  "#q-app > div.container > div.q-layout > div.layout-view > div > div.acoesHome > div:nth-child(2) > div > div.q-item.q-item-division.relative-position > div.q-item-main.q-item-section > div";
const btnAgendar =
  "#q-app > div.container > div.q-layout > div.layout-view > div > div.acoesHome > div.q-collapsible.q-item-division.relative-position.titleacord.user.q-collapsible-opened.q-collapsible-cursor-pointer > div > div:nth-child(2) > div > div > p:nth-child(2) > button > div.q-btn-inner.row.col.items-center.q-popup--skip.justify-center";
const btnGru =
  "#geral > div:nth-child(4) > div.q-tabs.flex.no-wrap.overflow-hidden.q-tabs-position-top.q-tabs-inverted > div.q-tabs-head.row.q-tabs-align-left.text-primary > div.q-tabs-scroller.row.no-wrap > div:nth-child(2)";
const btnAddServico =
  "#geral > div:nth-child(4) > div:nth-child(2) > button > div.q-btn-inner.row.col.items-center.q-popup--skip.justify-center";
const capitaniaBrasilia =
  "body > div.q-popover.scroll.column.no-wrap.animate-popup-down > div > div:nth-child(12) > div > div";
const capitaniaGo =
  "body > div.q-popover.scroll.column.no-wrap.animate-popup-down > div > div:nth-child(13) > div > div";
const btnTermos = "#checkorientacoes > div > i:nth-child(3)";
const firstProximoBtn =
  "#q-app > div.container > div.q-layout > div.layout-view > div > div.borderGeralEtapa > div.stepper-box > div.bottom.only-next > div > div.stepper-button.next";
const selectCPF = "#tipoSelectInteressado";
const btnSelecionarGrupo =
  "#geral > div:nth-child(4) > div.q-tabs.flex.no-wrap.overflow-hidden.q-tabs-position-top.q-tabs-inverted > div.q-tabs-panes > div > div > div > div > div.tdfristserv.col-md-4.col-xs-12 > div > div > div > div.q-if.row.no-wrap.relative-position.q-select.groupSelect.q-if-has-label.q-if-focusable.q-if-inverted.q-if-has-content.bg-white.text-white > div.q-if-inner.col.column.q-popup--skip > div.row.no-wrap.relative-position > div.col.q-input-target.ellipsis.justify-start";
const btnEducacional =
  "body > div.q-popover.scroll.column.no-wrap.animate-popup-down > div > div:nth-child(3) > div > div";
const btnCardeneta =
  "#geral > div:nth-child(4) > div.q-tabs.flex.no-wrap.overflow-hidden.q-tabs-position-top.q-tabs-inverted > div.q-tabs-panes > div > div > div > div > div.col-md-7.col-xs-12 > div > div > div > div.q-if.row.no-wrap.relative-position.q-select.servSelectGroup.q-if-has-label.q-if-focusable.q-if-inverted.q-if-has-content.bg-white.text-white > div.q-if-inner.col.column.q-popup--skip > div.row.no-wrap.relative-position > div.col.q-input-target.ellipsis.justify-start";
const btnProximo =
  "#q-app > div.container > div.q-layout > div.layout-view > div > div.borderGeralEtapa > div.stepper-box > div.bottom > div > div.stepper-button.next";
const btnHoraManha =
  "#geral > div.row.gutter-xs > div.col-md-8 > div > div > div.row.xs-gutter > div:nth-child(1) > div > div > div > div > div > div:nth-child(1) > div > div > i.q-icon.q-radio-checked.cursor-pointer.absolute-full.material-icons";
const btnHoraTarde =
  "#geral > div.row.gutter-xs > div.col-md-8 > div > div > div.row.xs-gutter > div:nth-child(2) > div > div > div > div > div > div:nth-child(1) > div > div > i.q-icon.q-radio-unchecked.cursor-pointer.absolute-full.material-icons";
const btnDataDisponivel =
  "#geral > div.row.gutter-xs > div.col-md-4 > div > div > div.__vev_calendar-wrapper > div.cal-wrapper > div.cal-body > div.dates > div.item.event";

let cpf = "";

// adicionar delay
function waitFor(delay) {
  return new Promise((resolve) => setTimeout(resolve, delay));
}

for (let index = 2; index < 3; index++) {
  cpf = worksheet[`A${index}`].v;
  console.log({ cpf: cpf });
}

async function robo() {
  const browser = await puppeteer.launch({
    headless: false,
    args: ["--start-maximized"],
  });
  const page = await browser.newPage();
  await page.setViewport({ width: 1920, height: 1080 });

  await page.goto("https://sistemas.dpc.mar.mil.br/sisap/agendamento/#/");

  // esperar e clicar no campo
  await page.waitForSelector(selectOrgMilitar);
  await page.click(selectOrgMilitar);

  // esperar opcao
  await waitFor(300);
  await page.waitForSelector(capitaniaGo);
  await page.click(capitaniaGo);

  // btn agendamento
  await page.waitForSelector(btnAgendamento);
  await page.click(btnAgendamento);
  await waitFor(200);

  // clique para agendar

  await page.waitForSelector(btnAgendar);
  await page.click(btnAgendar);
  await waitFor(1000);
  // page.waitForNavigation({ waitUntil: "networkidle2" });

  if (
    (await page.$(
      "body > div.modal.fullscreen.row.minimized.flex-center > div.modal-content"
    )) !== null // se nao tiver horario em Goias
  ) {
    // recomecar e escolher Brasilia
    await page.goto("https://sistemas.dpc.mar.mil.br/sisap/agendamento/#/");
    await page.waitForSelector(selectOrgMilitar);
    await page.click(selectOrgMilitar);
    await waitFor(500);
    await page.waitForSelector(capitaniaBrasilia);
    await page.click(capitaniaBrasilia);
    await page.waitForSelector(btnAgendamento);
    await page.click(btnAgendamento);
    await waitFor(200);
    await page.waitForSelector(btnAgendar);
    await page.click(btnAgendar);
  }

  // concordo com os termos
  await page.waitForSelector(btnTermos);
  await page.click(btnTermos);

  // first proximo btn

  await page.waitForSelector(firstProximoBtn);
  await page.click(firstProximoBtn);

  // clicar no selecione CPF
  await page.waitForSelector(selectCPF);
  await page.click(selectCPF);

  // clicar no CPF
  await page.waitForSelector(
    "body > div.q-popover.scroll.column.no-wrap.animate-popup-down > div"
  );

  await page.keyboard.press("ArrowDown");
  await page.keyboard.press("Enter");
  await waitFor(100);

  // digitar CPF
  await page.type("#cpfcnpjInteressado", `--------${cpf}`, { delay: 10 });

  // proximo
  await page.click(
    "#q-app > div.container > div.q-layout > div.layout-view > div > div.borderGeralEtapa > div.stepper-box > div.bottom > div > div.stepper-button.next"
  );

  // SEM GRU
  await page.waitForSelector(btnGru);
  await page.click(btnGru);

  // add servico
  await page.waitForSelector(btnAddServico);
  await page.click(btnAddServico);

  // selecionar grupo
  await page.waitForSelector(btnSelecionarGrupo);
  await page.click(btnSelecionarGrupo);
  await waitFor(300);

  // educacional
  await page.waitForSelector(btnEducacional);
  await page.click(btnEducacional);

  // cardeneta
  await page.waitForSelector(btnCardeneta);
  await page.click(btnCardeneta);
  await waitFor(300);

  await page.waitForSelector(
    "body > div.q-popover.scroll.column.no-wrap.animate-popup-up > div"
  );

  await page.click(
    "body > div.q-popover.scroll.column.no-wrap.animate-popup-up > div > div:nth-child(4) > div > div"
  );

  // proximo
  await page.click(btnProximo);

  // calendario DIA
  await page.waitForSelector(
    "#geral > div.row.gutter-xs > div.col-md-4 > div > div"
  );

  await page.waitForSelector(btnDataDisponivel);

  await page.click(btnDataDisponivel);
  await waitFor(200);

  // HORA

  // if ((await page.$(btnHoraManha)) !== null) {
  //   await waitFor(900);
  //   await page.click(btnHoraManha);
  // }
  // else {
  //   await waitFor(200);
  //   await page.waitForSelector(btnHoraTarde);
  //   await page.click(btnHoraTarde);
  // }

  // await page.click(
  //   "#geral > div.row.gutter-xs > div.col-md-8 > div > div > div.row.xs-gutter > div:nth-child(1) > div > div > div > div > div > div > div > div > i.q-icon.q-radio-unchecked.cursor-pointer.absolute-full.material-icons"
  // );

  // await page.click(btnPRoximo);

  // termos
  // await page.waitForSelector("#checkorientacoes > div > i:nth-child(3)");
  // await page.click("#checkorientacoes > div > i:nth-child(3)");

  // await browser.close();
}
robo();
