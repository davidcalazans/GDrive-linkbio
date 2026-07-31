function doGet(e) {
  var action = e.parameter.action;
  if (action === "admin") {
    return getAdminPage();
  }
  return HtmlService.createHtmlOutput(getTreeView())
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  var senhaSalva = sheet.getRange("A2").getValue().toString().trim();
  var senhaDigitada = e.parameter.senha ? e.parameter.senha.toString().trim() : "";

  // Captura direta e segura do parâmetro 'action' independente de onde ele venha
  var acao = e.parameter.action ? e.parameter.action.toString().trim() : "";

  // 1. Rota de Login por senha mestre
  if (acao === "login") {
    if (senhaDigitada === senhaSalva) {
      return getPainelPage(sheet);
    } else {
      Utilities.sleep(10000); // Bloqueio de 10 segundos contra força bruta
      return getAdminPage(true);
    }
  }

  // 2. Ação de Logout/Sair
  if (acao === "logout") {
    var props = PropertiesService.getScriptProperties();
    props.deleteProperty("TOKEN_ATIVO");
    props.deleteProperty("HORA_LOGIN");
    return getAdminPage();
  }

  // 3. Ação: Salvar Alterações (Validado por Token e Tempo Limite)
  if (acao === "salvar") {
    var tokenEnviado = e.parameter.tokenSessao;
    var props = PropertiesService.getScriptProperties();
    var tokenValido = props.getProperty("TOKEN_ATIVO");
    var horaLoginStr = props.getProperty("HORA_LOGIN");

    // Verificação de tempo (30 minutos)
    if (horaLoginStr) {
      var horaLogin = parseInt(horaLoginStr, 10);
      if ((Date.now() - horaLogin) > 1800000) {
        props.deleteProperty("TOKEN_ATIVO");
        props.deleteProperty("HORA_LOGIN");
        return ContentService.createTextOutput("SESSAO_EXPIRADA").setMimeType(ContentService.MimeType.TEXT);
      }
    }

    // Validação estrita do Token de Sessão
    if (!tokenEnviado || tokenEnviado !== tokenValido) {
      return ContentService.createTextOutput("ERRO_AUTENTICACAO").setMimeType(ContentService.MimeType.TEXT);
    }

    // Validação do tamanho da frase-senha (Mínimo 20 caracteres)
    if (e.parameter.novasenha && e.parameter.novasenha.trim() !== "") {
      var novaSenhaLimpa = e.parameter.novasenha.trim();
      if (novaSenhaLimpa.length < 20) {
        return ContentService.createTextOutput("ERRO_SENHA_CURTA").setMimeType(ContentService.MimeType.TEXT);
      }
      sheet.getRange("A2").setValue(novaSenhaLimpa);
      props.deleteProperty("TOKEN_ATIVO");
      props.deleteProperty("HORA_LOGIN");
    }

    // Grava as informações atualizadas na planilha
    sheet.getRange("B2").setValue(e.parameter.nome);
    sheet.getRange("C2").setValue(e.parameter.foto);
    sheet.getRange("D2").setValue(e.parameter.fundo);
    sheet.getRange("E2").setValue(e.parameter.links);
    sheet.getRange("F2").setValue(e.parameter.descricao);
    sheet.getRange("G2").setValue(e.parameter.video);
    sheet.getRange("H2").setValue(e.parameter.facebookUrl);
    sheet.getRange("I2").setValue(e.parameter.instagramUrl);
    sheet.getRange("J2").setValue(e.parameter.tiktokUrl);
    sheet.getRange("K2").setValue(e.parameter.youtubeUrl);
    sheet.getRange("L2").setValue(e.parameter.spotifyUrl);
    sheet.getRange("M2").setValue(e.parameter.xtwitterUrl);
    sheet.getRange("N2").setValue(e.parameter.pinterestUrl);
    sheet.getRange("O2").setValue(e.parameter.linkedinUrl);
    sheet.getRange("P2").setValue(e.parameter.githubUrl);
    sheet.getRange("Q2").setValue(e.parameter.telegramUrl);
    sheet.getRange("R2").setValue(e.parameter.whatsappUrl);

    return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
  } // Fim do bloco "salvar"

  return getAdminPage();
} // Fim da função doPost

function verificarSenhaNativa(senhaDigitada) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  var senhaSalva = sheet.getRange("A2").getValue().toString().trim();
  
  // Retorna o resultado imediatamente para evitar o timeout do navegador
  return (senhaDigitada.toString().trim() === senhaSalva);
}

// TELA 1: Visão Pública da Árvore de Links (VERSÃO CORRIGIDA E PROTEGIDA)
function getTreeView() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  var nome = sheet.getRange("B2").getValue();
  var foto = sheet.getRange("C2").getValue().toString().trim();
  var fundo = sheet.getRange("D2").getValue().toString().trim();
  var linksRaw = sheet.getRange("E2").getValue().toString();
  var descricao = sheet.getRange("F2").getValue().toString();
  var videoEmbedCode = sheet.getRange("G2").getValue().toString().trim();
  let facebookUrl = sheet.getRange("H2").getValue();
  let instagramUrl = sheet.getRange("I2").getValue();
  let tiktokUrl = sheet.getRange("J2").getValue();
  let youtubeUrl = sheet.getRange("K2").getValue();
  let spotifyUrl = sheet.getRange("L2").getValue();
  let xtwitterUrl = sheet.getRange("M2").getValue();
  let pinterestUrl = sheet.getRange("N2").getValue();
  let linkedinUrl = sheet.getRange("O2").getValue();
  let githubUrl = sheet.getRange("P2").getValue();
  let telegramUrl = sheet.getRange("Q2").getValue();
  let whatsappUrl = sheet.getRange("R2").getValue();

  
  var bgStyle = (fundo.includes("http") || fundo.includes("drive.")) ? "background-image: url('" + fundo + "'); background-size: cover; background-position: center;" : "background-color: " + fundo + ";";
  
  let facebookHtml = '';
    if (facebookUrl && facebookUrl.toString().trim() !== "") {
      facebookHtml = `
        <a href="${facebookUrl.toString().trim()}" target="_blank" rel="noopener noreferrer">
          <img class="imagemicone" src="https://img.icons8.com/?size=100&id=118467&format=png&color=000000" alt="Facebook">
        </a>
      `;
      }

  let instagramHtml = '';
    if (instagramUrl && instagramUrl.toString().trim() !== "") {
      instagramHtml = `
        <a href="${instagramUrl.toString().trim()}" target="_blank" rel="noopener noreferrer">
          <img class="imagemicone" src="https://img.icons8.com/?size=100&id=32309&format=png&color=000000" alt="Instagram">
        </a>
      `;
      }

  let tiktokhtml = '';
    if (tiktokUrl && tiktokUrl.toString().trim() !== "") {
      tiktokhtml = `
        <a href="${tiktokUrl.toString().trim()}" target="_blank" rel="noopener noreferrer">
          <img class="imagemicone" src="https://img.icons8.com/?size=100&id=118638&format=png&color=000000" alt="Tiktok">
        </a>
      `;
      }

  let youtubeHtml = '';
    if (youtubeUrl && youtubeUrl.toString().trim() !== "") {
      youtubeHtml = `
        <a href="${youtubeUrl.toString().trim()}" target="_blank" rel="noopener noreferrer">
          <img class="imagemicone" src="https://img.icons8.com/?size=100&id=37326&format=png&color=000000" alt="YouTube">
        </a>
      `;
      }

  let spotifyHtml = '';
    if (spotifyUrl && spotifyUrl.toString().trim() !== "") {
      spotifyHtml = `
        <a href="${spotifyUrl.toString().trim()}" target="_blank" rel="noopener noreferrer">
          <img class="imagemicone" src="https://img.icons8.com/?size=100&id=11116&format=png&color=000000" alt="Spotify">
        </a>
      `;
      }

  let xtwitterHtml = '';
    if (xtwitterUrl && xtwitterUrl.toString().trim() !== "") {
      xtwitterHtml = `
        <a href="${xtwitterUrl.toString().trim()}" target="_blank" rel="noopener noreferrer">
          <img class="imagemicone" src="https://img.icons8.com/?size=100&id=phOKFKYpe00C&format=png&color=000000" alt="X/Twitter">
        </a>
      `;
      }

  let pinterestHtml = '';
    if (pinterestUrl && pinterestUrl.toString().trim() !== "") {
      pinterestHtml = `
        <a href="${pinterestUrl.toString().trim()}" target="_blank" rel="noopener noreferrer">
          <img class="imagemicone" src="https://img.icons8.com/?size=100&id=11101&format=png&color=000000" alt="Pinterest">
        </a>
      `;
      }

  let linkedinHtml = '';
    if (linkedinUrl && linkedinUrl.toString().trim() !== "") {
      linkedinHtml = `
        <a href="${linkedinUrl.toString().trim()}" target="_blank" rel="noopener noreferrer">
          <img class="imagemicone" src="https://img.icons8.com/?size=100&id=8808&format=png&color=000000" alt="LinkedIn">
        </a>
      `;
      }

  let githubHtml = '';
    if (githubUrl && githubUrl.toString().trim() !== "") {
      githubHtml = `
        <a href="${githubUrl.toString().trim()}" target="_blank" rel="noopener noreferrer">
          <img class="imagemicone" src="https://img.icons8.com/?size=100&id=12599&format=png&color=000000" alt="GitHub">
        </a>
      `;
      }

  let telegramHtml = '';
    if (telegramUrl && telegramUrl.toString().trim() !== "") {
      telegraHtml = `
        <a href="${telegramUrl.toString().trim()}" target="_blank" rel="noopener noreferrer">
          <img class="imagemicone" src="https://img.icons8.com/?size=100&id=F4ZPUh2Mk5tk&format=png&color=000000" alt="Telegram">
        </a>
      `;
      }

  let whatsappHtml = '';
    if (whatsappUrl && whatsappUrl.toString().trim() !== "") {
      whatsappHtml = `
        <a href="${whatsappUrl.toString().trim()}" target="_blank" rel="noopener noreferrer">
          <img class="imagemicone" src="https://img.icons8.com/?size=100&id=16733&format=png&color=000000" alt="WhatsApp">
        </a>
      `;
      }

  // RENDERIZADOR DE BOTÕES ORIGINAL RESTAURADO (INTACTO)
  var linksHtml = "";
  if (linksRaw) {
    var lista = linksRaw.split("|");
    lista.forEach(function(item) {
      var partes = item.split(",");
      if(partes.length >= 2) {
        // Pega o nome (antes da primeira vírgula) e a URL (depois da primeira vírgula)
        var nomeBotao = partes[0];
        var urlBotao = partes[1];
        linksHtml += "<a class='btn-link' href='" + urlBotao + "' target='_blank'>" + nomeBotao + "</a>";
      }
    });
  }

  // NOVO MOTOR DE VIDEO PROTEGIDO (SEM ALTERAR TEXTOS OU LINKS)
  var videoHtml = "";
  if (videoEmbedCode && videoEmbedCode.includes("<iframe")) {
    var isVertical = videoEmbedCode.includes("shorts") || videoEmbedCode.includes("instagram") || videoEmbedCode.includes("tiktok");
    var paddingRatio = isVertical ? "177.77%" : "56.25%"; 
    var maxWidthVideo = isVertical ? "340px" : "100%";

    videoHtml = `
      <div class="video-section" style="max-width: ${maxWidthVideo};">
        <div class="video-container" style="padding-top: ${paddingRatio};">
          ${videoEmbedCode}
        </div>
      </div>
    `;
  }

  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <!-- NOVO: Importação da fonte moderna 'Plus Jakarta Sans' do Google Fonts -->
  <link rel="preconnect" href="https://googleapis.com">
  <link rel="preconnect" href="https://gstatic.com" crossorigin>
  <link href="https://googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;700;800&display=swap" rel="stylesheet">

  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; -webkit-tap-highlight-color: transparent; }
    html, body { height: 100%; width: 100%; overflow-x: hidden; }
    body {
      /* Alterado: Aplicando a nova fonte em todo o documento */
      font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
      display: flex; flex-direction: column; align-items: center; justify-content: flex-start;
      min-height: 100vh; padding: 40px 6vw; color: white;
      ${bgStyle} background-attachment: scroll; background-position: center center; background-size: cover; background-repeat: no-repeat;
    }
    .imagemicone { width: 40px; height: 40px; transition: transform 0.2s; }
    .wrapper { width: 100%; max-width: 580px; display: flex; flex-direction: column; align-items: center; flex-grow: 1; }
    
    .profile-card {
      width: 100%;
      background: rgba(240, 240, 240, 0.30);
      backdrop-filter: blur(4px);
      -webkit-backdrop-filter: blur(4px);
      border-radius: 24px;
      padding: 24px 20px;
      display: flex;
      flex-direction: column;
      align-items: center;
      margin-bottom: 24px;
      box-shadow: 0 8px 32px rgba(0, 0, 0, 0.15);
    }

    .profile-img { width: 110px; height: 110px; object-fit: cover; border-radius: 24px; border: 3px solid rgba(255, 255, 255, 0.95); box-shadow: 0 4px 14px rgba(0, 0, 0, 0.15); flex-shrink: 0; }
    
    /* Ajustado: Peso da fonte (800) para um visual marcante e moderno */
    h1 { margin: 8px 0 8px 0; font-size: 2.1rem; font-weight: 600; text-align: center; width: 100%; letter-spacing: -0.03em; color: #FFFFFF; }
    
    .profile-desc { font-size: 1.15rem; font-weight: 500; text-align: center; width: 100%; max-width: 420px; color: #FFFFFF; opacity: 0.85; line-height: 1.5; white-space: pre-wrap; }
    
    .links-container { width: 100%; display: flex; flex-direction: column; gap: 16px; margin-bottom: 24px; margin-top: 15px; }
.btn-link { 
      display: flex; 
      align-items: center; 
      justify-content: center; 
      width: 100%; 
      padding: 18px 20px; 
      text-align: center; 
      text-decoration: none; 
      color: #000000;
      
      /* Fonte de alta legibilidade para telas */
      font-family: 'Inter', sans-serif;
      font-weight: 300; 
      font-size: 1.15rem; 
      letter-spacing: -0.01em;
      
      background: rgba(240, 240, 240, 0.80); 
      backdrop-filter: blur(4px);
      -webkit-backdrop-filter: blur(4px);
      border-radius: 24px; 
      box-shadow: 0 4px 14px rgba(0, 0, 0, 0.15); 
      transition: transform 0.1s, background-color 0.1s; 
      word-break: break-word; 
    }
    .btn-link:active { transform: scale(0.97); background: #e5e5e5; }
    @media (hover: hover) { .btn-link:hover { transform: scale(1.015); background: #f7f7f7; } }
    .video-section { width: 100%; margin: 16px auto 24px auto; border-radius: 20px; overflow: hidden; box-shadow: 0 8px 24px rgba(0,0,0,0.5); background: #000; }
    .video-container { position: relative; width: 100%; height: 0; }
    .video-container iframe { position: absolute; top: 0; left: 0; width: 100% !important; height: 100% !important; border: 0; }
    .admin-gate { margin-top: auto; padding-top: 50px; font-size: 0.8rem; opacity: 0.5; text-decoration: none; color: white; text-shadow: 0 1px 3px rgba(0, 0, 0, 0.5); letter-spacing: 0.03em; }
  </style>
  <script>
    document.addEventListener("DOMContentLoaded", function () {
      const icones = document.querySelectorAll("img.imagemicone");
      icones.forEach(function (img) {
        img.style.transition = "transform 0.2s ease-in-out";
        img.style.cursor = "pointer";
        img.addEventListener("mouseover", function () { this.style.transform = "scale(1.3)"; });
        img.addEventListener("mouseout", function () { this.style.transform = "scale(1.0)"; });
      });
    });
  </script>
</head>
<body>
  <div class="wrapper">
    <div class="profile-card">
      <img class="profile-img" src="${foto}" alt="Profile">
      <h1>${nome}</h1>
      <div class="social-icons-container" style="display: flex; justify-content: center; align-items: center; gap: 14px;flex-wrap: wrap;">
        ${facebookHtml} ${instagramHtml} ${tiktokhtml} ${youtubeHtml} ${spotifyHtml} ${xtwitterHtml} ${pinterestHtml} ${linkedinHtml} ${githubHtml} ${telegramHtml} ${whatsappHtml}
      </div>
      <p class="profile-desc">${descricao}</p>
    </div>
    
    ${videoHtml}
    
    <div class="links-container">${linksHtml}</div>
    <a class="admin-gate" href="${ScriptApp.getService().getUrl()}?action=admin" target="_blank">Gerenciar Árvore</a>
  </div>
</body>
</html>
  `;
}

function getAdminPage(comErro) {
  var template = HtmlService.createTemplateFromFile('AdminTemplate');
  template.urlApp = ScriptApp.getService().getUrl();
  template.erroLogin = comErro || false;
  
  // CORREÇÃO: Força o Apps Script a injetar as regras mobile direto no topo da janela do navegador
  return template.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

function getPainelPage(sheet) {
  var nome = sheet.getRange("B2").getValue();
  var foto = sheet.getRange("C2").getValue();
  var fundo = sheet.getRange("D2").getValue();
  var links = sheet.getRange("E2").getValue();
  var descricao = sheet.getRange("F2").getValue(); 
  var video = sheet.getRange("G2").getValue();
  var facebookUrl = sheet.getRange("H2").getValue();
  var instagramUrl = sheet.getRange("I2").getValue();
  var tiktokUrl = sheet.getRange("J2").getValue();
  var youtubeUrl = sheet.getRange("K2").getValue();
  var spotifyUrl = sheet.getRange("L2").getValue();
  var xtwitterUrl = sheet.getRange("M2").getValue();
  var pinterestUrl = sheet.getRange("N2").getValue();
  var linkedinUrl = sheet.getRange("O2").getValue();
  var githubUrl = sheet.getRange("P2").getValue();
  var telegramUrl = sheet.getRange("Q2").getValue();
  var whatsappUrl = sheet.getRange("R2").getValue();

  var timestampAtual = Date.now().toString();
  var tokenSessao = Utilities.base64EncodeWebSafe(sheet.getRange("A2").getValue() + "_" + timestampAtual);

  var props = PropertiesService.getScriptProperties();
  props.setProperty("TOKEN_ATIVO", tokenSessao);
  props.setProperty("HORA_LOGIN", timestampAtual);

  var template = HtmlService.createTemplateFromFile('PainelTemplate');
  
  template.tokenSessao = tokenSessao;
  template.nome = nome;
  template.foto = foto;
  template.fundo = fundo;
  template.links = links;
  template.descricao = descricao;
  template.video = video;
  template.facebookUrl = facebookUrl;
  template.instagramUrl = instagramUrl;
  template.tiktokUrl = tiktokUrl;
  template.youtubeUrl = youtubeUrl;
  template.spotifyUrl = spotifyUrl;
  template.xtwitterUrl = xtwitterUrl;
  template.pinterestUrl = pinterestUrl;
  template.linkedinUrl = linkedinUrl;
  template.githubUrl = githubUrl;
  template.telegramUrl = telegramUrl;
  template.whatsappUrl = whatsappUrl;

  // CORREÇÃO: Força a injeção mobile no contêiner do painel administrativo
  return template.evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}
