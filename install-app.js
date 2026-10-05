(() => {
  'use strict';
  let installPrompt = null;
  let installed = false;
  const standalone = window.matchMedia('(display-mode: standalone)');
  const isInstalled = () => installed || standalone.matches || navigator.standalone === true;
  const buttons = () => document.querySelectorAll('[data-install-app]');
  function refresh() {
    buttons().forEach(button => { button.hidden = isInstalled(); });
  }
  window.addEventListener('beforeinstallprompt', event => {
    event.preventDefault();
    installPrompt = event;
    refresh();
  });
  window.addEventListener('appinstalled', () => {
    installed = true;
    installPrompt = null;
    document.getElementById('installAppDialog')?.close();
    refresh();
  });
  standalone.addEventListener('change', refresh);

  function showHelp() {
    let dialog = document.getElementById('installAppDialog');
    if (!dialog) {
      dialog = document.createElement('dialog');
      dialog.id = 'installAppDialog';
      dialog.className = 'install-app-dialog';
      dialog.setAttribute('aria-labelledby', 'installAppTitle');
      const appleMobile = /iPad|iPhone|iPod/.test(navigator.userAgent) ||
        (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
      const android = /Android/.test(navigator.userAgent);
      const mac = /Mac/.test(navigator.platform);
      const steps = appleMobile
        ? ['Ouvre cette page dans Safari.', 'Appuie sur Partager (le carré avec une flèche vers le haut), puis « Sur l’écran d’accueil ».', 'Si proposé, active « Ouvrir comme app », puis appuie sur « Ajouter ».']
        : android
          ? ['Ouvre cette page dans Chrome.', 'Ouvre le menu ⋮, puis « Ajouter à l’écran d’accueil » ou « Installer l’application ».', 'Confirme pour retrouver Marion parmi tes applications.']
          : mac
            ? ['Dans Safari, ouvre le menu Fichier, puis « Ajouter au Dock ».', 'Dans Chrome ou Edge, utilise l’option d’installation dans la barre d’adresse ou le menu du navigateur.']
            : ['Ouvre cette page dans Chrome ou Microsoft Edge.', 'Utilise l’icône d’installation dans la barre d’adresse, ou l’option « Installer » / « Applications » dans le menu du navigateur.', 'Si l’application est déjà installée, ouvre-la depuis tes applications.'];
      dialog.innerHTML = '<h2 id="installAppTitle">Installer Tournée Marion</h2><p>Retrouve les tournées et le calendrier depuis une icône sur ton appareil.</p><ol></ol><p class="install-app-note">Une connexion Internet reste nécessaire pour utiliser l’application.</p><form method="dialog"><button autofocus>Fermer</button></form>';
      for (const step of steps) {
        const li = document.createElement('li');
        li.textContent = step;
        dialog.querySelector('ol').append(li);
      }
      document.body.append(dialog);
    }
    if (!dialog.open) dialog.showModal();
  }
  async function install(event) {
    if (isInstalled()) { refresh(); return; }
    if (!installPrompt) { showHelp(); return; }
    const prompt = installPrompt;
    installPrompt = null;
    const button = event.currentTarget;
    button.disabled = true;
    try {
      await prompt.prompt();
      await prompt.userChoice;
    } catch {
      showHelp();
    } finally {
      button.disabled = false;
      refresh();
    }
  }
  function init() {
    buttons().forEach(button => button.addEventListener('click', install));
    refresh();
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, {once:true});
  else init();
})();
