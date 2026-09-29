/* ===== Login Protegido - JavaScript ===== */
/* Edite apenas o objeto USUARIOS abaixo */
/* Este arquivo serve para TODAS as páginas */

const USUARIOS = {
  'rai': '1234',
  'Rai': '1234',
  'mari': '1234',
  'Mari': '1234',
  'MARI': '1234',
  'mariana':'1234',
  'MARIANA':'1234',
  'izaque': '1234'
  // Adicione mais usuários assim:
  // 'nome': 'senha',
};

// =====================================================
// Não precisa mexer daqui para baixo
// =====================================================

(function () {
  const loginOverlay = document.getElementById('loginOverlay');
  const loginUser = document.getElementById('loginUser');
  const loginPass = document.getElementById('loginPass');
  const loginBtn = document.getElementById('loginBtn');
  const loginError = document.getElementById('loginError');
  const mainApp = document.getElementById('mainApp');
  const btnSair = document.getElementById('btnSair');

  if (!loginOverlay || !mainApp) {
    console.error('Login: elementos #loginOverlay ou #mainApp não encontrados.');
    return;
  }

  // Já logado nesta sessão? Mostra o conteúdo direto
  if (sessionStorage.getItem('logado') === 'true') {
    loginOverlay.style.display = 'none';
    mainApp.style.display = 'block';
  }

  function tentarLogin() {
    const user = loginUser.value.trim();
    const pass = loginPass.value;

    if (USUARIOS[user] && USUARIOS[user] === pass) {
      sessionStorage.setItem('logado', 'true');
      loginOverlay.style.display = 'none';
      mainApp.style.display = 'block';
      loginError.style.display = 'none';
    } else {
      loginError.style.display = 'block';
      loginPass.value = '';
      loginPass.focus();
    }
  }

  function sair() {
    sessionStorage.removeItem('logado');
    location.reload();
  }

  if (loginBtn) loginBtn.addEventListener('click', tentarLogin);

  if (loginPass) {
    loginPass.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') tentarLogin();
    });
  }

  if (loginUser) {
    loginUser.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') loginPass && loginPass.focus();
    });
  }

  if (btnSair) {
    btnSair.addEventListener('click', sair);
  }

  // Bloquear F12 e teclas de inspeção (fraco, só dificulta um pouco)
  document.addEventListener('keydown', function (e) {
    if (e.key === 'F12' || e.keyCode === 123) {
      e.preventDefault();
      return false;
    }
    if (e.ctrlKey && e.shiftKey && ['I', 'i', 'J', 'j', 'C', 'c'].includes(e.key)) {
      e.preventDefault();
      return false;
    }
    if (e.ctrlKey && (e.key === 'u' || e.key === 'U')) {
      e.preventDefault();
      return false;
    }
  });
})();
