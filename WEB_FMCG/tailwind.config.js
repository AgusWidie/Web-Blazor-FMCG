/** @type {import('tailwindcss').Config} */
module.exports = {
    content: [
        "./Pages/**/*.razor",
        "./Shared/**/*.razor",
        "./Layout/**/*.razor",
        "./Components/**/*.razor",
        "./wwwroot/**/*.html"
    ],
    theme: {
        extend: {
            animation: {
                'spin-slow': 'spin 3s linear infinite',
            }
        }
    },
    plugins: [],

}
