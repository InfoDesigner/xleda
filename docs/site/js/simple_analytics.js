(function() {
    // 1. Create and inject the main JavaScript tracker
    var s = document.createElement('script');
    s.async = true;
    s.defer = true;
    s.src = 'http://scripts.simpleanalyticscdn.com/';
    document.head.appendChild(s);

    // 2. Create and inject the noscript image fallback for non-JS users
    var ns = document.createElement('noscript');
    var img = document.createElement('img');
    img.src = 'http://scripts.simpleanalyticscdn.com/';
    img.alt = '';
    img.referrerPolicy = 'no-referrer-when-downgrade';
    ns.appendChild(img);
    document.body.appendChild(ns);
})();
