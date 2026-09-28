
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
  const force=document.getElementById('forca');
  const corpRow=document.getElementById('corpDataRow');
  const corpHint=document.getElementById('corpHint');
  const uf=document.getElementById('ufCorporacao');
  const rank=document.getElementById('postoGraduacao');
  if(!force||!corpRow||!uf||!rank)return;

  const common=[
    ['Cel PM','Coronel PM'],['Ten-Cel PM','Tenente-Coronel PM'],['Maj PM','Major PM'],['Cap PM','Capitão PM'],
    ['1º Ten PM','1º Tenente PM'],['2º Ten PM','2º Tenente PM'],['Asp Of PM','Aspirante a Oficial PM'],['Al Of PM','Aluno-Oficial PM'],
    ['ST PM','Subtenente PM'],['1º Sgt PM','1º Sargento PM'],['2º Sgt PM','2º Sargento PM'],['3º Sgt PM','3º Sargento PM'],
    ['Cb PM','Cabo PM'],['Sd PM','Soldado PM']
  ];
  const commonBm=[
    ['Cel BM','Coronel BM'],['Ten-Cel BM','Tenente-Coronel BM'],['Maj BM','Major BM'],['Cap BM','Capitão BM'],
    ['1º Ten BM','1º Tenente BM'],['2º Ten BM','2º Tenente BM'],['Asp Of BM','Aspirante a Oficial BM'],['Al Of BM','Aluno-Oficial BM'],
    ['ST BM','Subtenente BM'],['1º Sgt BM','1º Sargento BM'],['2º Sgt BM','2º Sargento BM'],['3º Sgt BM','3º Sargento BM'],
    ['Cb BM','Cabo BM'],['Sd BM','Soldado BM']
  ];
  const army=[
    ['Gen Ex EB','General de Exército'],['Gen Div EB','General de Divisão'],['Gen Bda EB','General de Brigada'],
    ['Cel EB','Coronel'],['Ten-Cel EB','Tenente-Coronel'],['Maj EB','Major'],['Cap EB','Capitão'],['1º Ten EB','1º Tenente'],['2º Ten EB','2º Tenente'],
    ['Asp Of EB','Aspirante a Oficial'],['Cad EB','Cadete'],['S Ten EB','Subtenente'],['1º Sgt EB','1º Sargento'],['2º Sgt EB','2º Sargento'],['3º Sgt EB','3º Sargento'],['Cb EB','Cabo'],['Sd EB','Soldado']
  ];
  const navy=[
    ['Alte MB','Almirante'],['Alte Esq MB','Almirante de Esquadra'],['V Alte MB','Vice-Almirante'],['C Alte MB','Contra-Almirante'],
    ['CMG MB','Capitão de Mar e Guerra'],['CF MB','Capitão de Fragata'],['CC MB','Capitão de Corveta'],['CT MB','Capitão-Tenente'],
    ['1º Ten MB','1º Tenente'],['2º Ten MB','2º Tenente'],['GM MB','Guarda-Marinha'],['Asp Of MB','Aspirante a Oficial'],
    ['SO MB','Suboficial'],['1º SG MB','1º Sargento'],['2º SG MB','2º Sargento'],['3º SG MB','3º Sargento'],['CB MB','Cabo'],['MN MB','Marinheiro']
  ];
  const air=[
    ['Mar FAB','Marechal do Ar'],['Ten Brig Ar','Tenente-Brigadeiro do Ar'],['Maj Brig Ar','Major-Brigadeiro do Ar'],['Brig Ar','Brigadeiro do Ar'],
    ['Cel FAB','Coronel'],['Ten-Cel FAB','Tenente-Coronel'],['Maj FAB','Major'],['Cap FAB','Capitão'],['1º Ten FAB','1º Tenente'],['2º Ten FAB','2º Tenente'],
    ['Asp Of FAB','Aspirante a Oficial'],['Cad FAB','Cadete'],['SO FAB','Suboficial'],['1S FAB','1º Sargento'],['2S FAB','2º Sargento'],['3S FAB','3º Sargento'],['Cb FAB','Cabo'],['S1 FAB','Soldado de 1ª Classe'],['S2 FAB','Soldado de 2ª Classe']
  ];

  function options(list){
    rank.innerHTML='<option value="">Selecione</option>'+list.map(x=>'<option value="'+x[0]+'">'+x[1]+'</option>').join('');
  }
  function update(){
    const v=force.value;
    const needsRank=['Brigada Militar – RS','Polícia Militar – outro Estado','Corpo de Bombeiros Militar – RS','Corpo de Bombeiros Militar – outro Estado','Exército Brasileiro','Marinha do Brasil','Força Aérea Brasileira'].includes(v);
    const needsUf=['Polícia Militar – outro Estado','Corpo de Bombeiros Militar – outro Estado'].includes(v);
    corpRow.hidden=!needsRank;
    corpHint.hidden=!needsRank;
    rank.required=needsRank;
    uf.required=needsUf;
    uf.closest('label').style.display=needsUf?'block':'none';
    if(!needsRank){rank.value='';rank.innerHTML='<option value="">Selecione</option>';uf.value='';}
    else if(v==='Brigada Militar – RS')options(common);
    else if(v.startsWith('Polícia Militar'))options(common);
    else if(v.startsWith('Corpo de Bombeiros'))options(commonBm);
    else if(v==='Exército Brasileiro')options(army);
    else if(v==='Marinha do Brasil')options(navy);
    else if(v==='Força Aérea Brasileira')options(air);
  }
  force.addEventListener('change',update);
  update();
})();

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
