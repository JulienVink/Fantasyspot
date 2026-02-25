let startPos = null;
let watchId = null;
let player = null;
let hasTriggered = false;
const RADIUS_THRESHOLD = 250; // meters

const startBtn = document.getElementById('start-btn');
const statusDiv = document.getElementById('status');
const playerContainer = document.getElementById('player-container');

// YouTube IFrame API initialization
window.onYouTubeIframeAPIReady = function() {
    player = new YT.Player('player', {
        height: '240',
        width: '100%',
        playerVars: {
            'playsinline': 1,
            'autoplay': 0
        },
        events: {
            'onReady': onPlayerReady,
            'onError': onPlayerError
        }
    });
}

function onPlayerReady(event) {
    console.log("Player ready");
    statusDiv.innerText = "Ready. Click 'Start Tracking' to begin.";
}

function onPlayerError(event) {
    console.error("Player error:", event.data);
    if (hasTriggered) {
        statusDiv.innerHTML = 'Music player failed. <a href="https://music.youtube.com/search?q=Lazy+J+Plays" target="_blank">Click here to open YouTube Music</a>';
    }
}

function triggerMusic() {
    if (hasTriggered) return;
    hasTriggered = true;

    statusDiv.innerText = "Zone left! Starting music...";
    statusDiv.style.color = "red";

    playerContainer.style.display = "block";

    try {
        // Attempting to use the search-to-play feature
        player.loadPlaylist({
            list: 'Lazy J Plays',
            listType: 'search',
            index: 0,
            suggestedQuality: 'small'
        });
        player.playVideo();
    } catch (e) {
        onPlayerError(e);
    }
}

function calculateDistance(lat1, lon1, lat2, lon2) {
    const R = 6371e3; // Earth radius in meters
    const φ1 = lat1 * Math.PI / 180;
    const φ2 = lat2 * Math.PI / 180;
    const Δφ = (lat2 - lat1) * Math.PI / 180;
    const Δλ = (lon2 - lon1) * Math.PI / 180;

    const a = Math.sin(Δφ / 2) * Math.sin(Δφ / 2) +
              Math.cos(φ1) * Math.cos(φ2) *
              Math.sin(Δλ / 2) * Math.sin(Δλ / 2);
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));

    return R * c;
}

function success(pos) {
    const crd = pos.coords;

    if (!startPos) {
        startPos = crd;
        statusDiv.innerText = "Center set. Watching for 250m departure...";
        startBtn.disabled = true;
        console.log("Starting position:", startPos);
        return;
    }

    const dist = calculateDistance(startPos.latitude, startPos.longitude, crd.latitude, crd.longitude);
    console.log("Current distance:", dist.toFixed(2), "m");

    if (dist > RADIUS_THRESHOLD) {
        triggerMusic();
        if (watchId) {
            navigator.geolocation.clearWatch(watchId);
            watchId = null;
        }
    } else {
        statusDiv.innerText = "Inside zone. Distance: " + dist.toFixed(0) + "m";
    }
}

function error(err) {
    console.warn("ERROR(" + err.code + "): " + err.message);
    statusDiv.innerText = "Error getting location. Please enable GPS.";
}

startBtn.addEventListener('click', () => {
    if (!navigator.geolocation) {
        statusDiv.innerText = "Geolocation is not supported by your browser.";
        return;
    }

    statusDiv.innerText = "Requesting location...";

    // Start watching position
    watchId = navigator.geolocation.watchPosition(success, error, {
        enableHighAccuracy: true,
        timeout: 10000,
        maximumAge: 0
    });

    // Pre-load/Cue the player if possible to handle autoplay restrictions
    if (player && typeof player.cuePlaylist === 'function') {
        player.cuePlaylist({
            list: 'Lazy J Plays',
            listType: 'search',
            index: 0
        });
    }
});
