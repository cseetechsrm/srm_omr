window.addEventListener('DOMContentLoaded', async (event) => {
    const welcomeSlide = document.getElementById('welcomeSlide');
    const loginPage = document.getElementById('loginPage');

    // Helper function to simulate delay
    const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

    // Wait for 3 seconds, then transition the slides
    await delay(3000);
    welcomeSlide.classList.add('hide');

    // Wait for the transition to complete (1 second), then switch to login page
    await delay(1000);
    welcomeSlide.style.display = 'none'; // Hide welcome slide
    loginPage.classList.add('show'); // Show login page
});
