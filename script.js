
(function(){
  const form = document.getElementById('registrationForm');
  const message = document.getElementById('formMessage');
  function show(text, type){
    message.textContent = text;
    message.className = 'message show ' + type;
    message.scrollIntoView({behavior:'smooth', block:'center'});
  }

  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    const consent = form.querySelector('[name="consentimento"]');
    if (!consent.checked) return show('É necessário concordar com o registro dos dados.', 'error');

    const data = Object.fromEntries(new FormData(form).entries());
    data.timestamp = new Date().toISOString();
    data.evento = '38º Seminário do Instituto de Pesquisa da Brigada Militar';

    const endpoint = window.SEMINARIO_CONFIG?.webAppUrl || '';
    const button = form.querySelector('button');
    const oldText = button.textContent;

    if (!endpoint) {
      const list = JSON.parse(localStorage.getItem('ipbm_inscricoes_demo') || '[]');
      list.push(data);
      localStorage.setItem('ipbm_inscricoes_demo', JSON.stringify(list));
      show('Inscrição registrada no modo demonstração. Configure o Google Apps Script em config.js para receber inscrições reais.', 'success');
      form.reset();
      return;
    }

    button.disabled = true;
    button.textContent = 'Enviando...';
    try {
      await fetch(endpoint, {
        method: 'POST',
        mode: 'no-cors',
        headers: {'Content-Type':'text/plain;charset=utf-8'},
        body: JSON.stringify(data)
      });
      show('Inscrição enviada com sucesso. Confira seu e-mail para as próximas orientações.', 'success');
      form.reset();
    } catch (err) {
      show('Não foi possível enviar a inscrição. Revise a configuração do endereço do formulário e tente novamente.', 'error');
    } finally {
      button.disabled = false;
      button.textContent = oldText;
    }
  });
})();
