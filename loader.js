async function loadFragment(targetId, url) {
    const target = document.getElementById(targetId);
    if (!target) {
        throw new Error(`Container ${targetId} não encontrado.`);
    }
    const response = await fetch(url, { cache: 'no-cache' });
    if (!response.ok) {
        throw new Error(`Falha ao carregar ${url}: ${response.status}`);
    }
    target.innerHTML = await response.text();
}

function loadScriptSequential(src) {
    return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = src;
        script.onload = () => resolve();
        script.onerror = () => reject(new Error(`Falha ao carregar ${src}`));
        document.body.appendChild(script);
    });
}

async function bootstrapApp() {
    try {
        await loadFragment('loginViewContainer', 'views/login.html');
        await loadFragment('appViewContainer', 'views/app.html');

        await loadScriptSequential('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js');
        await loadScriptSequential('https://www.gstatic.com/firebasejs/10.12.2/firebase-app-compat.js');
        await loadScriptSequential('https://www.gstatic.com/firebasejs/10.12.2/firebase-auth-compat.js');
        await loadScriptSequential('https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore-compat.js');
        await loadScriptSequential('script.js');
        if (typeof window.startApp === 'function') {
            window.startApp();
        } else {
            throw new Error('Inicializador startApp não encontrado.');
        }
    } catch (error) {
        console.error(error);
        document.body.innerHTML = '<p style="padding:16px;font-family:Arial, sans-serif;">Erro ao carregar o sistema. Atualize a página.</p>';
    }
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', bootstrapApp);
} else {
    bootstrapApp();
}
