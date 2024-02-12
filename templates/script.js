document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('feedbackForm');

    form.addEventListener('submit', function(event) {
        event.preventDefault();

        const formData = new FormData(form);

        fetch('/submit-feedback', {
            method: 'POST',
            body: formData
        })
        .then(response => {
            if (response.ok) {
                return response.text();
            } else {
                throw new Error('Failed to submit feedback');
            }
        })
        .then(data => {
            console.log(data); 
        })
        .catch(error => {
            console.error('Error:', error); 
        });
    });
});
