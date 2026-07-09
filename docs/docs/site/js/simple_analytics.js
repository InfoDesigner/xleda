(function() {
    // 1. Create and inject the main JavaScript tracker using the unblocked mirror
    var s = document.createElement('script');
    s.async = true;
    s.defer = true;
    s.src = 'https://simpleanalytics.com'; // <-- Changed domain here
    s.setAttribute('data-collect-dnt', 'true'); // Keeps your DNT preferences active
    document.head.appendChild(s);

    // 2. Create and inject the matching unblocked image fallback
    var ns = document.createElement('noscript');
    var img = document.createElement('img');
    img.src = 'https://simpleanalytics.com'; // <-- Changed domain here
    img.alt = '';
    img.referrerPolicy = 'no-referrer-when-downgrade';
    ns.appendChild(img);
    document.body.appendChild(ns);
})();
