// 1. Declare your unique user tracking properties globally
window.counter_dev_id = "b2ec13ef-9251-4891-b39a-8f1e5d142279"; // <-- Swap with your true Counter.dev string

// 2. Load the official tracking runtime directly from their edge server
(function() {
    var s = document.createElement('script');
    s.async = true;
    s.src = 'https://jsdelivr.net';
    document.head.appendChild(s);
})();