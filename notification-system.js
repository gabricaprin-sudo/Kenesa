// ============================================================
// NOTIFICATION SYSTEM — Admin Dashboard Integration
// Displays notifications from Firestore 'notifications' collection
// Add this to the END of app.js (before the closing /* === End of App === */)
// ============================================================

const NotificationSystem = {
  container: null,
  currentNotif: null,
  unsub: null,
  dismissedNotifs: new Set(),

  init() {
    // Load dismissed notifications from localStorage
    try {
      const dismissed = JSON.parse(localStorage.getItem('dismissedNotifs') || '[]');
      dismissed.forEach(id => this.dismissedNotifs.add(id));
    } catch (e) { console.warn('Failed to load dismissed notifications:', e); }

    // Create notification container
    this.container = document.createElement('div');
    this.container.id = 'notificationBanner';
    this.container.style.cssText = `
      position: fixed;
      top: calc(76px + env(safe-area-inset-top, 0px));
      left: 50%;
      transform: translateX(-50%);
      width: calc(100% - 32px);
      max-width: 480px;
      z-index: 1000;
      transition: all 0.4s cubic-bezier(0.34, 1.56, 0.64, 1);
      pointer-events: none;
    `;
    document.body.appendChild(this.container);

    // Add animation styles
    if (!document.getElementById('notifAnimStyles')) {
      const styleEl = document.createElement('style');
      styleEl.id = 'notifAnimStyles';
      styleEl.textContent = `
        @keyframes notifSlideIn {
          from { opacity: 0; transform: translateY(-20px) scale(0.95); }
          to { opacity: 1; transform: translateY(0) scale(1); }
        }
        @keyframes notifSlideOut {
          from { opacity: 1; transform: translateY(0) scale(1); }
          to { opacity: 0; transform: translateY(-20px) scale(0.95); }
        }
      `;
      document.head.appendChild(styleEl);
    }

    // Listen for notifications
    this.listenForNotifications();
  },

  listenForNotifications() {
    if (!firebaseReady || !db) {
      console.warn('Firebase not ready, retrying notifications in 3s...');
      setTimeout(() => this.listenForNotifications(), 3000);
      return;
    }

    try {
      const q = FB.query(
        FB.collection(db, 'notifications'),
        FB.where('active', '==', true),
        FB.orderBy('createdAt', 'desc')
      );

      this.unsub = FB.onSnapshot(q, (snap) => {
        const notifs = [];
        snap.forEach(doc => notifs.push({ id: doc.id, ...doc.data() }));

        // Filter out dismissed notifications
        const activeNotifs = notifs.filter(n => !this.dismissedNotifs.has(n.id));

        if (activeNotifs.length > 0) {
          // Show the most recent active notification
          this.showNotification(activeNotifs[0]);
        } else {
          this.hideNotification();
        }
      }, err => {
        console.error('Notification listener error:', err);
      });
    } catch (e) {
      console.error('Failed to init notification listener:', e);
    }
  },

  showNotification(notif) {
    if (this.currentNotif === notif.id) return;
    this.currentNotif = notif.id;

    const typeColors = {
      info: { 
        bg: 'linear-gradient(135deg, #3498db 0%, #2980b9 100%)', 
        icon: '&#128161;',
        borderColor: '#3498db'
      },
      success: { 
        bg: 'linear-gradient(135deg, #2ecc71 0%, #27ae60 100%)', 
        icon: '&#9989;',
        borderColor: '#2ecc71'
      },
      warning: { 
        bg: 'linear-gradient(135deg, #f39c12 0%, #d35400 100%)', 
        icon: '&#9888;',
        borderColor: '#f39c12'
      },
      error: { 
        bg: 'linear-gradient(135deg, #e74c3c 0%, #c0392b 100%)', 
        icon: '&#10060;',
        borderColor: '#e74c3c'
      },
      feature: { 
        bg: 'linear-gradient(135deg, #9b59b6 0%, #8e44ad 100%)', 
        icon: '&#127881;',
        borderColor: '#9b59b6'
      }
    };

    const style = typeColors[notif.type] || typeColors.info;

    const dismissBtn = notif.dismissible ? 
      `<button onclick="NotificationSystem.dismiss('${notif.id}')" style="
        background: rgba(255,255,255,0.2);
        border: none;
        color: white;
        width: 28px;
        height: 28px;
        border-radius: 50%;
        cursor: pointer;
        font-size: 16px;
        display: flex;
        align-items: center;
        justify-content: center;
        flex-shrink: 0;
        transition: background 0.2s;
        pointer-events: all;
      " onmouseover="this.style.background='rgba(255,255,255,0.35)'" onmouseout="this.style.background='rgba(255,255,255,0.2)'">&#10005;</button>` : '';

    const link = notif.link ? 
      `<a href="${notif.link}" target="_blank" style="
        color: rgba(255,255,255,0.9);
        text-decoration: underline;
        font-size: 13px;
        margin-top: 6px;
        display: inline-block;
        pointer-events: all;
      ">المزيد &#10132;</a>` : '';

    this.container.innerHTML = `
      <div style="
        background: ${style.bg};
        color: white;
        border-radius: 16px;
        padding: 16px 20px;
        box-shadow: 0 8px 32px rgba(0,0,0,0.25);
        display: flex;
        align-items: flex-start;
        gap: 14px;
        pointer-events: all;
        animation: notifSlideIn 0.5s cubic-bezier(0.34, 1.56, 0.64, 1);
        backdrop-filter: blur(10px);
        -webkit-backdrop-filter: blur(10px);
        border: 2px solid ${style.borderColor};
      ">
        <div style="font-size: 28px; flex-shrink: 0; line-height: 1;">${style.icon}</div>
        <div style="flex: 1; min-width: 0;">
          <div style="font-size: 15px; font-weight: 700; margin-bottom: 4px; line-height: 1.3;">${esc(notif.title)}</div>
          <div style="font-size: 13px; opacity: 0.9; line-height: 1.5;">${esc(notif.body)}</div>
          ${link}
        </div>
        ${dismissBtn}
      </div>
    `;
  },

  hideNotification() {
    this.currentNotif = null;
    if (this.container) {
      this.container.innerHTML = '';
    }
  },

  dismiss(id) {
    // Mark as dismissed in localStorage
    this.dismissedNotifs.add(id);
    try {
      const dismissed = Array.from(this.dismissedNotifs);
      localStorage.setItem('dismissedNotifs', JSON.stringify(dismissed));
    } catch (e) { console.warn('Failed to save dismissed notification:', e); }

    // Animate out
    const banner = this.container.querySelector('div');
    if (banner) {
      banner.style.animation = 'notifSlideOut 0.3s ease forwards';
      setTimeout(() => this.hideNotification(), 300);
    }
  },

  destroy() {
    if (this.unsub) this.unsub();
    if (this.container) this.container.remove();
  }
};

// Auto-initialize after app is ready
// Hook into the existing bootstrap function
const _originalBootstrap = bootstrap;
bootstrap = async function() {
  await _originalBootstrap();
  // Initialize notifications after a short delay to ensure Firebase is ready
  setTimeout(() => {
    if (typeof NotificationSystem !== 'undefined' && firebaseReady) {
      NotificationSystem.init();
      console.log('✅ Notification system initialized');
    }
  }, 2000);
};
