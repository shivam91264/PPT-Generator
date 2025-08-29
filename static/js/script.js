document.getElementById('pptForm').addEventListener('submit', function(e) {
    e.preventDefault();
    const form = e.target;
    const loading = document.getElementById('loading');
    const message = document.getElementById('message');
    loading.style.display = 'block';
    message.textContent = '';
    const formData = new FormData(form);

    fetch(form.action, {
        method: 'POST',
        body: formData
    })
    .then(async response => {
        loading.style.display = 'none';
        if (response.ok) {
            const blob = await response.blob();
            const dlLink = document.createElement('a');
            dlLink.href = window.URL.createObjectURL(blob);
            dlLink.download = "generated_presentation.pptx";
            document.body.appendChild(dlLink);
            dlLink.click();
            dlLink.remove();
            message.textContent = "Your presentation is ready!";
        } else {
            message.textContent = "Error generating presentation. Please check your inputs.";
        }
    })
    .catch(() => {
        loading.style.display = 'none';
        message.textContent = "Network or server error. Please try again.";
    });
});
