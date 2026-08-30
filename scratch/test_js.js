
        // Instant Theme Init to prevent screen flicker (FOUC)
        (function() {
            const saved = localStorage.getItem('eic_theme') || 'dark';
            document.documentElement.setAttribute('data-theme', saved);
        })();
    