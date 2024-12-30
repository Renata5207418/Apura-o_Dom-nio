// index.js

// Alternar visibilidade da senha
function togglePassword() {
    const passwordField = document.getElementById("senha");
    const toggleIcon = document.querySelector(".toggle-password i");

    if (passwordField.type === "password") {
        passwordField.type = "text";
        toggleIcon.classList.replace("fa-eye-slash", "fa-eye");
    } else {
        passwordField.type = "password";
        toggleIcon.classList.replace("fa-eye", "fa-eye-slash");
    }
}

// Enviar formulário com dados
document.addEventListener("DOMContentLoaded", () => {
    const form = document.getElementById("unified-form");

    form.addEventListener("submit", async (event) => {
        event.preventDefault();

        const formData = new FormData();
        formData.append("usuario", document.getElementById("usuario").value);
        formData.append("senha", document.getElementById("senha").value);
        formData.append("arquivo_empresas", document.getElementById("arquivo_empresas").files[0]);

        try {
            const response = await fetch("/upload-and-login", {
                method: "POST",
                body: formData,
            });

            const result = await response.json();
            if (response.ok) {
                alert(result.message);
            } else {
                alert("Erro: " + result.error);
            }
        } catch (error) {
            alert("Erro ao enviar os dados: " + error);
        }
    });
});

// Atualizar progresso a cada 10 segundos
async function atualizarProgresso() {
    const response = await fetch("/progresso");
    const mensagens = await response.json();

    const progressContainer = document.getElementById("progress-container");
    progressContainer.innerHTML = "";
    mensagens.forEach(msg => {
        const p = document.createElement("p");
        p.textContent = msg;
        progressContainer.appendChild(p);
    });
}

setInterval(atualizarProgresso, 10000);
