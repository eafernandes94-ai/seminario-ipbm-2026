
(function(){
  const modal = document.getElementById('certificateModal');
  const frame = document.getElementById('certificateFrame');
  const openButton = document.querySelector('.certificate-open');
  const certUrl = 'https://script.google.com/macros/s/AKfycbwH_9PQK6OtZUQtozM8WDAusJJ69V8j8ylQwUlTJ6mZYOd3l6tJv0O1dx_eAXXxqiXvFg/exec?cert=1';

  function openCertificate(){
    frame.src = certUrl;
    modal.classList.add('open');
    modal.setAttribute('aria-hidden','false');
    document.body.classList.add('modal-open');
  }
  function closeCertificate(){
    modal.classList.remove('open');
    modal.setAttribute('aria-hidden','true');
    frame.src = 'about:blank';
    document.body.classList.remove('modal-open');
  }
  if(openButton) openButton.addEventListener('click', openCertificate);
  document.querySelectorAll('[data-close-cert]').forEach((el)=>el.addEventListener('click', closeCertificate));
  document.addEventListener('keydown',(e)=>{
    if(e.key==='Escape' && modal.classList.contains('open')) closeCertificate();
  });
})();


;
document.querySelectorAll('.cv-toggle').forEach((button) => {
  button.addEventListener('click', () => {
    const panel = document.getElementById(button.getAttribute('aria-controls'));
    const opening = panel.hasAttribute('hidden');
    if (opening) panel.removeAttribute('hidden');
    else panel.setAttribute('hidden', '');
    button.setAttribute('aria-expanded', String(opening));
    button.textContent = opening ? 'Ocultar currículo' : 'Ver currículo';
  });
});


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
