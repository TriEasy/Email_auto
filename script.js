// Set the max date for the activity date input to today
const today = new Date();
const maxDate = today.toISOString().split('T')[0];
if (document.getElementById('activityDate')) {
    document.getElementById('activityDate').setAttribute('max', maxDate);
}

// Handle form submission
document.getElementById('knowledgeForm').addEventListener('submit', function (e) {
    e.preventDefault();

    document.getElementById('successMessage').style.display = 'none';
    document.getElementById('loading').style.display = 'block';

    // Collect data from all form fields
    const formData = {
        department: document.getElementById('department').value,
        activityTopic: document.getElementById('activityTopic').value,
        activityType: document.getElementById('activityType').value,
        strategicGoalLevel1: document.getElementById('strategicGoalLevel1').value,
        strategicGoalLevel2: document.getElementById('strategicGoalLevel2').value,
        presenterCategory: document.getElementById('presenterCategory').value,
        activityDate: document.getElementById('activityDate').value,
        presenterName: document.getElementById('presenterName').value,
        attendanceResponsible: document.getElementById('attendanceResponsible').value,
        targetAudience: document.getElementById('targetAudience').value,
        attendeeCount: document.getElementById('attendeeCount').value,
        activityDuration: document.getElementById('activityDuration').value,
        contentLocation: document.getElementById('contentLocation').value,
    };

    // Send the data to the Python server
    fetch('/submit', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify(formData),
    })
    .then(response => {
        if (!response.ok) {
            throw new Error('Network response was not ok');
        }
        return response.json();
    })
    .then(data => {
        console.log('Success:', data);
        document.getElementById('loading').style.display = 'none';
        document.getElementById('successMessage').style.display = 'block';
        document.getElementById('successMessage').scrollIntoView({ behavior: 'smooth' });

        // Reset the form after 3 seconds
        setTimeout(function () {
            document.getElementById('knowledgeForm').reset();
            document.getElementById('successMessage').style.display = 'none';
            window.scrollTo({ top: 0, behavior: 'smooth' });
        }, 3000);
    })
    .catch((error) => {
        console.error('Error:', error);
        document.getElementById('loading').style.display = 'none';
        // Use a less intrusive error message
        const errorDiv = document.createElement('div');
        errorDiv.style.color = 'red';
        errorDiv.style.marginTop = '10px';
        errorDiv.textContent = 'An error occurred while submitting the form. Please check the server console and try again.';
        document.getElementById('knowledgeForm').appendChild(errorDiv);
    });
});

