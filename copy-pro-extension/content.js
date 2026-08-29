(function() {
  let lastTap = 0;

  // 🎨 Inject CSS styles for feedback ripple overlays
  const style = document.createElement('style');
  style.textContent = `
    .pw-ext-skip-overlay {
      position: absolute !important;
      top: 50% !important;
      transform: translateY(-50%) !important;
      width: 70px !important;
      height: 70px !important;
      background: rgba(0, 0, 0, 0.7) !important;
      border: 1px solid rgba(255, 255, 255, 0.2) !important;
      border-radius: 50% !important;
      display: flex !important;
      flex-direction: column !important;
      align-items: center !important;
      justify-content: center !important;
      color: white !important;
      font-family: system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif !important;
      font-size: 11px !important;
      font-weight: 800 !important;
      opacity: 0 !important;
      pointer-events: none !important;
      transition: all 0.2s cubic-bezier(0.175, 0.885, 0.32, 1.275) !important;
      z-index: 2147483647 !important;
    }
    .pw-ext-skip-overlay.active {
      opacity: 1 !important;
      transform: translateY(-50%) scale(1.15) !important;
    }
    .pw-ext-skip-overlay.left {
      left: 15% !important;
    }
    .pw-ext-skip-overlay.right {
      right: 15% !important;
    }
  `;
  document.documentElement.appendChild(style);

  function setupPlayerListeners(video) {
    if (video.dataset.pwExtRegistered) return;
    video.dataset.pwExtRegistered = "true";

    // Standard custom players wrap the video element inside a container
    const parent = video.parentElement || document.body;
    parent.style.position = 'relative';

    // Create Left Skip Feedback Indicator
    let leftOverlay = document.createElement('div');
    leftOverlay.className = 'pw-ext-skip-overlay left';
    leftOverlay.innerHTML = '<div style="font-size: 20px; font-weight: 800; margin-bottom: 2px;">◀◀</div><div>10s</div>';
    parent.appendChild(leftOverlay);

    // Create Right Skip Feedback Indicator
    let rightOverlay = document.createElement('div');
    rightOverlay.className = 'pw-ext-skip-overlay right';
    rightOverlay.innerHTML = '<div style="font-size: 20px; font-weight: 800; margin-bottom: 2px;">▶▶</div><div>10s</div>';
    parent.appendChild(rightOverlay);

    function showSkipFeedback(direction) {
      const overlay = (direction === 'left') ? leftOverlay : rightOverlay;
      overlay.classList.add('active');
      setTimeout(() => {
        overlay.classList.remove('active');
      }, 500);
    }

    // Intercept clicks/taps on the video container
    parent.addEventListener('click', function(e) {
      const rect = parent.getBoundingClientRect();
      const clickY = e.clientY - rect.top;
      const clickX = e.clientX - rect.left;
      
      // Ignore clicks near controls bar (bottom 20% height of player)
      if (clickY > rect.height * 0.8) return;

      const now = Date.now();
      const DOUBLE_TAP_DELAY = 300;
      if (now - lastTap < DOUBLE_TAP_DELAY) {
        // Intercept play/pause toggle of native custom player controls
        e.preventDefault();
        e.stopPropagation();

        if (clickX < rect.width / 2) {
          // Left Double Tap -> Rewind 10s
          video.currentTime = Math.max(0, video.currentTime - 10);
          showSkipFeedback('left');
        } else {
          // Right Double Tap -> Forward 10s
          video.currentTime = Math.min(video.duration || 0, video.currentTime + 10);
          showSkipFeedback('right');
        }
      }
      lastTap = now;
    }, true); // Use capture phase to intercept before custom player click events play/pause video
  }

  // Periodic detector to handle dynamic SPAs and lazy loaded video players
  setInterval(() => {
    const videos = document.querySelectorAll('video');
    videos.forEach(setupPlayerListeners);
  }, 1000);

  console.log("Copy Pro - PW Player Gesture Controller extension successfully loaded!");
})();
