<!doctype html>
<html lang="id">
 <head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover">
  <meta name="theme-color" content="#dc2626">
  <meta name="mobile-web-app-capable" content="yes">
  <meta name="apple-mobile-web-app-capable" content="yes">
  <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
  <meta name="apple-mobile-web-app-title" content="SNBT Attin">
  <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
  <meta name="format-detection" content="telephone=no">
  <title>SNBT Consultation - Bimbel Attin v1.4.9</title>
  <script src="/_sdk/element_sdk.js"></script>
  <script src="/_sdk/data_sdk.js"></script>
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
  <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            -webkit-tap-highlight-color: transparent;
        }

        html {
            height: 100%;
            width: 100%;
            overflow-x: hidden;
        }

        body {
            box-sizing: border-box;
            margin: 0 !important;
            padding: 0 !important;
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
            background: linear-gradient(135deg, #1a0000 0%, #2d0a0a 25%, #4a0000 50%, #2d0a0a 75%, #1a0000 100%) !important;
            background-attachment: fixed;
            height: 100vh;
            width: 100vw;
            overflow-x: hidden;
            overflow-y: hidden;
            -webkit-font-smoothing: antialiased;
            -moz-osx-font-smoothing: grayscale;
            font-size: 16px;
            touch-action: pan-y;
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            will-change: auto;
            contain: layout style paint;
        }

        .mobile-container {
            width: 100%;
            height: 100%;
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            overflow-x: hidden;
            overflow-y: auto;
            display: block;
            contain: layout style paint;
            will-change: auto;
            -webkit-overflow-scrolling: touch;
        }

        .glass-card {
            background: rgba(0, 0, 0, 0.6);
            backdrop-filter: blur(10px) saturate(180%);
            -webkit-backdrop-filter: blur(10px) saturate(180%);
            border-radius: 20px;
            border: 2px solid rgba(220, 38, 38, 0.4);
            box-shadow: 0 8px 32px rgba(220, 38, 38, 0.2), inset 0 0 20px rgba(220, 38, 38, 0.05);
            padding: 25px;
            contain: content;
        }

        .glass-button {
            background: linear-gradient(135deg, #dc2626, #991b1b);
            backdrop-filter: blur(10px);
            border: 2px solid rgba(239, 68, 68, 0.5);
            box-shadow: 0 4px 15px rgba(220, 38, 38, 0.4), 0 0 20px rgba(220, 38, 38, 0.2);
            transition: all 0.2s ease;
            padding: 16px 24px;
            border-radius: 12px;
            color: white;
            font-weight: 600;
            font-size: 16px;
            width: 100%;
            cursor: pointer;
            touch-action: manipulation;
            will-change: transform;
        }

        .glass-button:active {
            transform: scale(0.98);
            box-shadow: 0 2px 15px rgba(220, 38, 38, 0.8), 0 0 40px rgba(220, 38, 38, 0.5);
            background: linear-gradient(135deg, #b91c1c, #7f1d1d);
        }

        .glass-input {
            background: rgba(0, 0, 0, 0.5);
            backdrop-filter: blur(10px);
            border: 2px solid rgba(220, 38, 38, 0.3);
            border-radius: 12px;
            color: #fef2f2;
            padding: 12px 16px;
            width: 100%;
            transition: all 0.3s ease;
            font-weight: 600;
            font-size: 16px;
            outline: none;
            box-sizing: border-box;
            line-height: 1.5;
            height: auto;
            min-height: 44px;
            -webkit-user-select: text !important;
            -moz-user-select: text !important;
            user-select: text !important;
            -webkit-touch-callout: default;
            pointer-events: auto;
            touch-action: auto;
        }
        
        .glass-input::-webkit-input-placeholder {
            color: rgba(254, 242, 242, 0.5);
        }
        
        .glass-input::-moz-placeholder {
            color: rgba(254, 242, 242, 0.5);
        }

        .glass-input:focus {
            outline: none;
            background: rgba(0, 0, 0, 0.7);
            border-color: #dc2626;
            box-shadow: 0 0 25px rgba(220, 38, 38, 0.6), 0 0 15px rgba(220, 38, 38, 0.4);
        }

        .glass-input::placeholder {
            color: rgba(254, 242, 242, 0.5);
        }

        select.glass-input {
            appearance: none;
            background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='24' height='24' viewBox='0 0 24 24' fill='none' stroke='white' stroke-width='2'%3E%3Cpolyline points='6 9 12 15 18 9'%3E%3C/polyline%3E%3C/svg%3E");
            background-repeat: no-repeat;
            background-position: right 12px center;
            background-size: 20px;
            padding-right: 45px;
            cursor: pointer;
        }

        select.glass-input option {
            background: #667eea;
            color: white;
            padding: 12px;
        }

        .page-container {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            width: 100%;
            height: 100%;
            padding: 20px;
            padding-top: 140px;
            padding-bottom: 100px;
            display: none;
            overflow-x: hidden;
            overflow-y: auto;
            will-change: auto;
            backface-visibility: hidden;
            -webkit-backface-visibility: hidden;
            transform: translateZ(0);
            contain: layout style paint;
            -webkit-overflow-scrolling: touch;
        }

        .page-container.active {
            display: block !important;
            visibility: visible;
            opacity: 1;
            z-index: 10;
        }
        
        #page-1.page-container {
            padding-top: 20px;
            top: 0;
        }

        .floating-orb {
            position: fixed;
            border-radius: 50%;
            filter: blur(50px);
            opacity: 0.15;
            animation: float 20s ease-in-out infinite;
            pointer-events: none;
            z-index: 0;
            will-change: transform;
            backface-visibility: hidden;
            -webkit-backface-visibility: hidden;
        }

        @keyframes float {
            0%, 100% { transform: translate(0, 0); }
            50% { transform: translate(10px, -10px); }
        }

        .glow-text {
            text-shadow: 0 0 25px rgba(254, 242, 242, 0.95),
                         0 0 45px rgba(220, 38, 38, 0.8),
                         0 0 60px rgba(220, 38, 38, 0.5);
            font-weight: 700;
            color: #fef2f2;
        }

        .choice-card {
            cursor: pointer;
            transition: all 0.2s ease;
            padding: 25px;
            border-radius: 16px;
            position: relative;
            touch-action: manipulation;
            will-change: transform;
            transform: translateZ(0);
        }

        .choice-card:active {
            transform: scale(0.98);
        }

        .choice-card.selected {
            border: 3px solid #dc2626;
            box-shadow: 0 8px 30px rgba(220, 38, 38, 0.6), 0 0 20px rgba(220, 38, 38, 0.4);
        }

        .table-glass {
            background: rgba(0, 0, 0, 0.5);
            backdrop-filter: blur(15px);
            border-radius: 16px;
            overflow: hidden;
            border: 2px solid rgba(220, 38, 38, 0.3);
            width: 100%;
        }

        .table-glass th {
            background: rgba(220, 38, 38, 0.3);
            padding: 14px 10px;
            text-align: left;
            font-weight: 700;
            color: #fef2f2;
            font-size: 14px;
            border-bottom: 2px solid rgba(220, 38, 38, 0.4);
        }

        .table-glass td {
            padding: 12px 10px;
            color: rgba(254, 242, 242, 0.95);
            border-bottom: 1px solid rgba(220, 38, 38, 0.2);
            font-weight: 600;
            font-size: 13px;
        }

        .bottom-nav {
            position: fixed;
            bottom: 0;
            left: 0;
            right: 0;
            background: rgba(0, 0, 0, 0.95);
            backdrop-filter: blur(20px);
            border-top: 2px solid rgba(220, 38, 38, 0.3);
            display: flex;
            justify-content: space-around;
            padding: 12px 0;
            z-index: 150;
            box-shadow: 0 -4px 20px rgba(220, 38, 38, 0.2);
            height: 70px;
            align-items: center;
        }

        .nav-item {
            flex: 1;
            display: flex;
            flex-direction: column;
            align-items: center;
            color: rgba(254, 242, 242, 0.5);
            font-size: 11px;
            padding: 8px;
            cursor: pointer;
            transition: all 0.2s ease;
            touch-action: manipulation;
            will-change: color, text-shadow;
        }

        .nav-item.active {
            color: #fef2f2;
            text-shadow: 0 0 10px rgba(220, 38, 38, 0.8);
        }
        
        .nav-item.active svg {
            filter: drop-shadow(0 0 8px rgba(220, 38, 38, 0.8));
        }

        .nav-item svg {
            margin-bottom: 4px;
        }

        .header-bar {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            background: rgba(0, 0, 0, 0.95);
            backdrop-filter: blur(20px);
            padding: 16px 20px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            z-index: 100;
            box-shadow: 0 4px 20px rgba(220, 38, 38, 0.4), 0 2px 10px rgba(0, 0, 0, 0.5);
            border-bottom: 2px solid rgba(220, 38, 38, 0.3);
            height: 60px;
        }

        .header-title {
            color: white;
            font-size: 18px;
            font-weight: 700;
            text-shadow: 0 0 10px rgba(220, 38, 38, 0.4);
        }

        .back-button {
            color: white;
            padding: 8px;
            cursor: pointer;
            display: flex;
            align-items: center;
            touch-action: manipulation;
        }

        .toast {
            position: fixed;
            bottom: 90px;
            left: 50%;
            transform: translateX(-50%);
            background: rgba(0, 0, 0, 0.9);
            color: white;
            padding: 14px 24px;
            border-radius: 12px;
            z-index: 1000;
            animation: slideUp 0.3s ease;
            max-width: 90%;
            text-align: center;
        }

        @keyframes slideUp {
            from {
                opacity: 0;
                transform: translate(-50%, 20px);
            }
            to {
                opacity: 1;
                transform: translate(-50%, 0);
            }
        }

        .logo-container {
            width: 100px;
            height: 100px;
            margin: 0 auto 20px;
            border-radius: 50%;
            overflow: hidden;
            background: linear-gradient(135deg, rgba(220, 38, 38, 0.3), rgba(127, 29, 29, 0.3));
            padding: 8px;
            box-shadow: 0 10px 30px rgba(220, 38, 38, 0.6), 0 0 40px rgba(220, 38, 38, 0.3);
            border: 2px solid rgba(220, 38, 38, 0.6);
        }

        .logo-container img {
            width: 100%;
            height: 100%;
            object-fit: contain;
            border-radius: 50%;
        }

        .fab-button {
            position: fixed;
            bottom: 20px;
            right: 20px;
            width: 56px;
            height: 56px;
            border-radius: 50%;
            background: linear-gradient(135deg, #dc2626, #991b1b);
            box-shadow: 0 6px 20px rgba(220, 38, 38, 0.7), 0 0 30px rgba(220, 38, 38, 0.4);
            border: 2px solid rgba(239, 68, 68, 0.5);
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            cursor: pointer;
            z-index: 99;
            touch-action: manipulation;
            transition: transform 0.3s ease;
        }

        .fab-button:active {
            transform: scale(0.95);
        }

        .pg-modal-overlay {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(0, 0, 0, 0.8);
            backdrop-filter: blur(5px);
            z-index: 200;
            display: none;
            align-items: flex-end;
            animation: slideUp 0.3s ease;
        }

        .pg-modal-content {
            width: 100%;
            max-height: 85vh;
            background: rgba(0, 0, 0, 0.95);
            backdrop-filter: blur(20px);
            border-radius: 24px 24px 0 0;
            border: 2px solid rgba(220, 38, 38, 0.4);
            overflow-y: auto;
            padding: 0;
        }

        .pg-modal-header {
            position: sticky;
            top: 0;
            background: rgba(0, 0, 0, 0.95);
            backdrop-filter: blur(20px);
            padding: 20px;
            border-bottom: 2px solid rgba(220, 38, 38, 0.3);
            display: flex;
            justify-content: space-between;
            align-items: center;
            z-index: 10;
        }

        .pg-modal-title {
            color: #fef2f2;
            font-size: 18px;
            font-weight: 700;
            text-shadow: 0 0 10px rgba(220, 38, 38, 0.6);
        }

        .pg-modal-close {
            color: white;
            cursor: pointer;
            padding: 8px;
            display: flex;
            align-items: center;
            touch-action: manipulation;
        }

        .pg-modal-body {
            padding: 20px;
            padding-bottom: 30px;
        }

        @media print {
            body {
                background: white !important;
            }

            .floating-orb, .bottom-nav, .fab-button, .header-bar {
                display: none !important;
            }

            .page-container {
                display: block !important;
            }

            .glass-card {
                background: white !important;
                border: 2px solid #667eea !important;
                box-shadow: none !important;
            }

            * {
                color: #333 !important;
            }

            h1, h2, h3 {
                color: #667eea !important;
            }
        }

        .modal-overlay {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(0, 0, 0, 0.8);
            backdrop-filter: blur(5px);
            z-index: 300;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
        }

        .modal-content {
            max-width: 400px;
            width: 100%;
            max-height: 90%;
            overflow-y: auto;
        }

        .checkbox-custom {
            width: 20px;
            height: 20px;
            accent-color: #06d6a0;
            cursor: pointer;
        }

        .status-badge {
            display: inline-block;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 11px;
            font-weight: 700;
            margin-top: 4px;
        }

        .progress-circle {
            width: 70px;
            height: 70px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            background: conic-gradient(#06d6a0 0deg, #06d6a0 0deg, rgba(255,255,255,0.2) 0deg);
        }

        .progress-inner {
            width: 55px;
            height: 55px;
            border-radius: 50%;
            background: rgba(0, 0, 0, 0.3);
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: 700;
            color: white;
        }

        .content-wrapper {
            padding-top: 0;
            padding-bottom: 20px;
        }

        .loading-indicator {
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            background: rgba(0, 0, 0, 0.9);
            padding: 30px;
            border-radius: 20px;
            z-index: 9999;
            display: none;
        }

        .loading-spinner {
            width: 50px;
            height: 50px;
            border: 4px solid rgba(220, 38, 38, 0.3);
            border-top-color: #dc2626;
            border-radius: 50%;
            animation: spin 1s linear infinite;
        }

        @keyframes spin {
            to { transform: rotate(360deg); }
        }

        .admin-badge {
            display: inline-block;
            background: linear-gradient(135deg, #f5c857, #f093fb);
            color: white;
            padding: 4px 10px;
            border-radius: 20px;
            font-size: 11px;
            font-weight: 700;
            margin-left: 8px;
            text-transform: uppercase;
            letter-spacing: 0.3px;
        }

        .user-item {
            display: flex;
            justify-content: space-between;
            align-items: center;
            background: rgba(255,255,255,0.05);
            padding: 12px 16px;
            border-radius: 12px;
            margin-bottom: 10px;
            border: 1px solid rgba(220, 38, 38, 0.2);
        }

        .user-info {
            flex: 1;
        }

        .user-username {
            color: #fef2f2;
            font-weight: 600;
            font-size: 14px;
            margin-bottom: 3px;
        }

        .user-password {
            color: rgba(254, 242, 242, 0.6);
            font-size: 12px;
            font-family: monospace;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .password-text {
            letter-spacing: 2px;
        }

        .toggle-password-btn {
            background: rgba(220, 38, 38, 0.3);
            border: 1px solid rgba(220, 38, 38, 0.5);
            color: rgba(254, 242, 242, 0.8);
            padding: 2px 6px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 10px;
            transition: all 0.2s ease;
            font-weight: 600;
        }

        .toggle-password-btn:hover {
            background: rgba(220, 38, 38, 0.5);
        }

        .toggle-password-btn:active {
            transform: scale(0.95);
        }

        .delete-user-btn {
            background: linear-gradient(135deg, #f5576c, #f093fb);
            border: none;
            color: white;
            padding: 6px 12px;
            border-radius: 8px;
            cursor: pointer;
            font-size: 12px;
            font-weight: 600;
            transition: all 0.2s ease;
        }

        .delete-user-btn:active {
            transform: scale(0.95);
        }

        .marquee-container {
            overflow: hidden;
            width: 100%;
            max-width: 350px;
            margin: 0 auto;
        }

        .marquee-text {
            display: inline-block;
            animation: marquee 8s linear infinite;
            white-space: nowrap;
            padding-left: 100%;
        }

        @keyframes marquee {
            0% {
                transform: translateX(0);
            }
            100% {
                transform: translateX(-100%);
            }
        }
    </style>
  <script src="https://cdn.tailwindcss.com/3.4.17" type="text/javascript"></script>
  <script src="/_sdk/data_sdk.js" type="text/javascript"></script>
  <script src="https://cdn.jsdelivr.net/npm/lucide@0.263.0/dist/umd/lucide.min.js" type="text/javascript"></script>
 </head>
 <body>
  <div id="loading-indicator" class="loading-indicator">
   <div class="loading-spinner"></div>
  </div>
  <div class="mobile-container">
   <div class="floating-orb" style="width: 300px; height: 300px; background: linear-gradient(135deg, #7f1d1d, #dc2626); top: -100px; left: -100px;"></div>
   <div class="floating-orb" style="width: 250px; height: 250px; background: linear-gradient(135deg, #dc2626, #991b1b); bottom: -100px; right: -100px;"></div><!-- Page 1: Login -->
   <div id="page-1" class="page-container active" style="padding-top: 20px;">
    <div class="content-wrapper">
     <div style="text-align: center; padding: 40px 20px 30px;">
      <div class="logo-container"><img src="https://iili.io/KsKbLep.png" alt="Logo Bimbel Attin">
      </div>
      <h1 class="glow-text" style="font-size: 24px; font-weight: 700; margin-bottom: 10px; line-height: 1.3;">
       <div class="marquee-container">
        <span class="marquee-text">Konsultasi Jurusan &amp; Universitas</span>
       </div></h1>
      <h2 style="font-size: 20px; font-weight: 600; margin-bottom: 15px; color: white; opacity: 1;">Bimbel Attin</h2>
      <p style="color: white; opacity: 0.9; font-size: 14px; line-height: 1.6; max-width: 350px; margin: 0 auto;">Sistem rekomendasi program studi berdasarkan skor Try Out</p>
     </div>
     <div class="glass-card" style="margin: 0 20px;">
      <form id="login-form">
       <div style="margin-bottom: 20px;"><label for="username-input" style="display: block; margin-bottom: 10px; font-size: 14px; color: #fef2f2; font-weight: 600;">Username</label> <input type="text" class="glass-input" placeholder="Masukkan username" id="username-input" name="username" autocomplete="username" required>
       </div>
       <div style="margin-bottom: 25px;"><label for="password-input" style="display: block; margin-bottom: 10px; font-size: 14px; color: #fef2f2; font-weight: 600;">Password</label> <input type="password" class="glass-input" placeholder="Masukkan password" id="password-input" name="password" autocomplete="current-password" required>
       </div><button type="submit" class="glass-button" style="width: 100%;">
        <svg width="20" height="20" viewbox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="display: inline-block; margin-right: 8px; vertical-align: middle;"><path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4" /> <polyline points="10 17 15 12 10 7" /> <line x1="15" y1="12" x2="3" y2="12" />
        </svg> Login </button>
      </form>
     </div>
     <div style="text-align: center; padding: 30px 20px 20px; margin-top: 25px;">
      <div style="background: rgba(0, 0, 0, 0.4); backdrop-filter: blur(10px); border-radius: 12px; padding: 16px; border: 1px solid rgba(220, 38, 38, 0.3); margin: 0 20px;">
       <p style="color: #fef2f2; opacity: 0.7; font-size: 13px; margin-bottom: 6px; font-weight: 600;">© 2025 Bimbel Attin</p>
       <p style="color: #fef2f2; opacity: 0.6; font-size: 11px; margin-bottom: 4px;">Sistem Rekomendasi SNBT</p>
       <p style="color: #fef2f2; opacity: 0.5; font-size: 10px;">Version 1.4.9</p>
      </div>
     </div>
    </div>
   </div><!-- Page 2: Upload Data -->
   <div id="page-2" class="page-container">
    <div class="header-bar">
     <div class="back-button" onclick="showPage('page-1')">
      <svg width="24" height="24" viewbox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M19 12H5M12 19l-7-7 7-7" />
      </svg>
     </div>
     <div class="header-title">
      Upload Master Data
     </div>
     <div style="width: 24px;"></div>
    </div>
    <div class="content-wrapper">
     <div style="text-align: center; margin-bottom: 25px;">
      <div style="width: 70px; height: 70px; background: linear-gradient(135deg, #4facfe, #00f2fe); border-radius: 16px; margin: 0 auto 15px; display: flex; align-items: center; justify-content: center; box-shadow: 0 8px 20px rgba(79, 172, 254, 0.4);">
       <svg width="35" height="35" viewbox="0 0 24 24" fill="none" stroke="white" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /> <polyline points="17 8 12 3 7 8" /> <line x1="12" y1="3" x2="12" y2="15" />
       </svg>
      </div>
      <h2 class="text-white text-xl font-bold mb-2">Upload Data Excel</h2>
      <p class="text-white" style="opacity: 0.8; font-size: 14px;">Format: .xlsx / .csv</p>
     </div><input type="file" id="file-upload" accept=".xlsx,.csv,.xls" style="display: none;"> <button class="glass-button" onclick="document.getElementById('file-upload').click()" style="margin-bottom: 15px;">
      <svg width="20" height="20" viewbox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="display: inline-block; margin-right: 8px; vertical-align: middle;"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /> <polyline points="17 8 12 3 7 8" /> <line x1="12" y1="3" x2="12" y2="15" />
      </svg> Pilih File Excel </button>
     <p class="text-white text-center" style="opacity: 0.7; font-size: 13px;" id="file-name-display">Belum ada file dipilih</p>
     <div id="preview-card" class="glass-card" style="margin-top: 25px; display: none;">
      <h3 class="text-white font-bold mb-4" style="font-size: 18px;">Statistik Data</h3>
      <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px; margin-bottom: 20px;">
       <div style="background: rgba(220, 38, 38, 0.2); border: 2px solid rgba(220, 38, 38, 0.4); border-radius: 12px; padding: 15px; text-align: center;">
        <div style="font-size: 28px; font-weight: bold; color: #fef2f2; margin-bottom: 5px;" id="stat-rows">
         0
        </div>
        <div style="font-size: 12px; color: #fef2f2; opacity: 0.8; font-weight: 600;">
         Total Data
        </div>
       </div>
       <div style="background: rgba(220, 38, 38, 0.2); border: 2px solid rgba(220, 38, 38, 0.4); border-radius: 12px; padding: 15px; text-align: center;">
        <div style="font-size: 28px; font-weight: bold; color: #fef2f2; margin-bottom: 5px;" id="stat-universities">
         0
        </div>
        <div style="font-size: 12px; color: #fef2f2; opacity: 0.8; font-weight: 600;">
         Universitas
        </div>
       </div>
       <div style="background: rgba(220, 38, 38, 0.2); border: 2px solid rgba(220, 38, 38, 0.4); border-radius: 12px; padding: 15px; text-align: center;">
        <div style="font-size: 28px; font-weight: bold; color: #fef2f2; margin-bottom: 5px;" id="stat-majors">
         0
        </div>
        <div style="font-size: 12px; color: #fef2f2; opacity: 0.8; font-weight: 600;">
         Kategori Jurusan
        </div>
       </div>
       <div style="background: rgba(220, 38, 38, 0.2); border: 2px solid rgba(220, 38, 38, 0.4); border-radius: 12px; padding: 15px; text-align: center;">
        <div style="font-size: 28px; font-weight: bold; color: #fef2f2; margin-bottom: 5px;" id="stat-levels">
         0
        </div>
        <div style="font-size: 12px; color: #fef2f2; opacity: 0.8; font-weight: 600;">
         Jenjang
        </div>
       </div>
      </div>
      <h4 class="text-white font-bold mb-3" style="font-size: 15px;">Preview Data (5 teratas)</h4>
      <div style="overflow-x: auto; margin: 0 -25px; padding: 0 25px;">
       <table class="table-glass" style="font-size: 12px;">
        <thead>
         <tr>
          <th>Universitas</th>
          <th>Prodi</th>
          <th>PG</th>
         </tr>
        </thead>
        <tbody id="preview-tbody">
        </tbody>
       </table>
      </div>
     </div><button class="glass-button" onclick="showPage('page-3')" style="margin-top: 25px;"> Lanjutkan </button>
    </div>
   </div><!-- Page 3: Input Data Siswa -->
   <div id="page-3" class="page-container">
    <div class="header-bar">
     <div class="back-button" onclick="showPage('page-2')">
      <svg width="24" height="24" viewbox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M19 12H5M12 19l-7-7 7-7" />
      </svg>
     </div>
     <div class="header-title">
      Input Data Siswa
     </div>
     <div style="width: 24px;"></div>
    </div>
    <div class="content-wrapper">
     <div style="text-align: center; margin-bottom: 25px;">
      <div style="width: 70px; height: 70px; background: linear-gradient(135deg, #f093fb, #f5576c); border-radius: 16px; margin: 0 auto 15px; display: flex; align-items: center; justify-content: center; box-shadow: 0 8px 20px rgba(240, 147, 251, 0.4);">
       <svg width="35" height="35" viewbox="0 0 24 24" fill="none" stroke="white" stroke-width="2"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2" /> <circle cx="12" cy="7" r="4" />
       </svg>
      </div>
      <h2 class="text-white text-xl font-bold mb-2">Data Siswa</h2>
     </div>
     <div class="glass-card">
      <div style="margin-bottom: 20px;"><label for="student-name" class="text-white block mb-2" style="font-size: 14px;">Nama Lengkap</label> <input type="text" class="glass-input" placeholder="Masukkan nama lengkap" id="student-name">
      </div>
      <div style="margin-bottom: 20px;"><label for="student-school" class="text-white block mb-2" style="font-size: 14px;">Asal Sekolah</label> <input type="text" class="glass-input" placeholder="Nama sekolah" id="student-school">
      </div>
      <div style="margin-bottom: 25px;"><label for="student-score" class="text-white block mb-2" style="font-size: 14px;">Skor Try Out</label> <input type="number" step="0.001" class="glass-input" placeholder="Contoh: 650.123" id="student-score">
      </div><button class="glass-button" onclick="processStudentData()"> Lanjut ke Pemilihan </button>
     </div>
    </div>
   </div><!-- Page 4: Pilih Program Studi -->
   <div id="page-4" class="page-container">
    <div class="header-bar">
     <div class="back-button" onclick="showPage('page-3')">
      <svg width="24" height="24" viewbox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M19 12H5M12 19l-7-7 7-7" />
      </svg>
     </div>
     <div class="header-title">
      Pilih Program Studi
     </div>
     <div style="width: 24px;"></div>
    </div>
    <div class="content-wrapper">
     <div class="glass-card" style="text-align: center; margin-bottom: 20px;">
      <p class="text-white" style="opacity: 0.8; font-size: 14px; margin-bottom: 8px;">Skor Try Out Anda</p>
      <p class="text-white text-3xl font-bold" id="display-score">650</p>
     </div>
     <div style="display: grid; gap: 15px; margin-bottom: 20px;">
      <div id="choice-card-1" class="glass-card choice-card" style="background: linear-gradient(135deg, rgba(79, 172, 254, 0.2), rgba(0, 242, 254, 0.2));" onclick="selectChoice(1)">
       <div style="display: flex; justify-content: space-between; align-items: center;">
        <div style="flex: 1;">
         <div style="background: rgba(79, 172, 254, 0.4); padding: 6px 14px; border-radius: 20px; display: inline-block; margin-bottom: 10px;"><span class="text-white font-bold" style="font-size: 12px;">PILIHAN 1</span>
         </div>
         <p class="text-white" style="font-size: 14px; opacity: 0.9;">S1 • Bebas (Tanpa Batasan Skor)</p>
        </div>
        <div id="check-1" style="width: 28px; height: 28px; border-radius: 50%; background: rgba(255,255,255,0.3); display: none; align-items: center; justify-content: center;">
         <svg width="16" height="16" viewbox="0 0 24 24" fill="none" stroke="white" stroke-width="3"><polyline points="20 6 9 17 4 12" />
         </svg>
        </div>
       </div>
      </div>
      <div id="choice-card-2" class="glass-card choice-card" style="background: linear-gradient(135deg, rgba(118, 75, 162, 0.3), rgba(102, 126, 234, 0.3));" onclick="selectChoice(2)">
       <div style="display: flex; justify-content: space-between; align-items: center;">
        <div style="flex: 1;">
         <div style="background: rgba(118, 75, 162, 0.4); padding: 6px 14px; border-radius: 20px; display: inline-block; margin-bottom: 10px;"><span class="text-white font-bold" style="font-size: 12px;">PILIHAN 2</span>
         </div>
         <p class="text-white" style="font-size: 14px; opacity: 0.9;">S1 &amp; D4 • PG ≤ Skor TO</p>
        </div>
        <div id="check-2" style="width: 28px; height: 28px; border-radius: 50%; background: rgba(255,255,255,0.3); display: none; align-items: center; justify-content: center;">
         <svg width="16" height="16" viewbox="0 0 24 24" fill="none" stroke="white" stroke-width="3"><polyline points="20 6 9 17 4 12" />
         </svg>
        </div>
       </div>
      </div>
      <div id="choice-card-3" class="glass-card choice-card" style="background: linear-gradient(135deg, rgba(67, 97, 238, 0.3), rgba(72, 85, 201, 0.3));" onclick="selectChoice(3)">
       <div style="display: flex; justify-content: space-between; align-items: center;">
        <div style="flex: 1;">
         <div style="background: rgba(67, 97, 238, 0.4); padding: 6px 14px; border-radius: 20px; display: inline-block; margin-bottom: 10px;"><span class="text-white font-bold" style="font-size: 12px;">PILIHAN 3</span>
         </div>
         <p class="text-white" style="font-size: 14px; opacity: 0.9;">D4 &amp; D3 • PG ≤ (TO - 50)</p>
        </div>
        <div id="check-3" style="width: 28px; height: 28px; border-radius: 50%; background: rgba(255,255,255,0.3); display: none; align-items: center; justify-content: center;">
         <svg width="16" height="16" viewbox="0 0 24 24" fill="none" stroke="white" stroke-width="3"><polyline points="20 6 9 17 4 12" />
         </svg>
        </div>
       </div>
      </div>
      <div id="choice-card-4" class="glass-card choice-card" style="background: linear-gradient(135deg, rgba(6, 214, 160, 0.3), rgba(17, 153, 142, 0.3));" onclick="selectChoice(4)">
       <div style="display: flex; justify-content: space-between; align-items: center;">
        <div style="flex: 1;">
         <div style="background: rgba(6, 214, 160, 0.4); padding: 6px 14px; border-radius: 20px; display: inline-block; margin-bottom: 10px;"><span class="text-white font-bold" style="font-size: 12px;">PILIHAN 4</span>
         </div>
         <p class="text-white" style="font-size: 14px; opacity: 0.9;">D3 • PG ≤ (TO - 100)</p>
        </div>
        <div id="check-4" style="width: 28px; height: 28px; border-radius: 50%; background: rgba(255,255,255,0.3); display: none; align-items: center; justify-content: center;">
         <svg width="16" height="16" viewbox="0 0 24 24" fill="none" stroke="white" stroke-width="3"><polyline points="20 6 9 17 4 12" />
         </svg>
        </div>
       </div>
      </div>
     </div>
     <p class="text-white text-center" style="opacity: 0.8; font-size: 14px; margin-bottom: 15px;" id="choices-status">Pilih program studi untuk setiap pilihan</p><button id="analyze-button" class="glass-button" onclick="goToAnalysis()" style="display: none; background: linear-gradient(135deg, #06d6a0, #11998e);">
      <svg width="20" height="20" viewbox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="display: inline-block; margin-right: 8px; vertical-align: middle;"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12" />
      </svg> Lihat Analisis Peluang </button>
    </div>
   </div><!-- Page 5: List Program Studi -->
   <div id="page-5" class="page-container">
    <div class="header-bar">
     <div class="back-button" onclick="showPage('page-4')">
      <svg width="24" height="24" viewbox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M19 12H5M12 19l-7-7 7-7" />
      </svg>
     </div>
     <div class="header-title">
      Pilih Program Studi
     </div>
     <div style="width: 24px;"></div>
    </div>
    <div class="content-wrapper">
     <div class="glass-card" style="margin-bottom: 20px;">
      <div style="margin-bottom: 15px;"><input type="text" id="search-prodi" class="glass-input" placeholder="Cari jurusan..." oninput="applyFilters()">
      </div>
      <div style="display: grid; gap: 12px;"><select id="filter-university" class="glass-input" onchange="applyFilters()"> <option value="">Semua Universitas</option> </select> <select id="filter-level" class="glass-input" onchange="applyFilters()"> <option value="">Semua Jenjang</option> </select> <select id="filter-stream" class="glass-input" onchange="applyFilters()"> <option value="">Semua Prodi</option> <option value="IPA">IPA</option> <option value="IPS">IPS</option> </select>
      </div>
     </div><button class="glass-button" onclick="confirmSelection()" style="margin-bottom: 20px;">
      <svg width="20" height="20" viewbox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="display: inline-block; margin-right: 8px; vertical-align: middle;"><polyline points="20 6 9 17 4 12" />
      </svg> Konfirmasi Pilihan (<span id="selected-count">0</span>) </button>
     <p class="text-white text-center" style="opacity: 0.8; font-size: 13px; margin-bottom: 15px;" id="result-count">0 program studi yang disarankan</p>
     <div id="programs-list" style="display: grid; gap: 12px;">
     </div>
    </div>
   </div><!-- Page Admin: User Management -->
   <div id="page-admin" class="page-container">
    <div class="header-bar">
     <div class="back-button" onclick="adminLogout()">
      <svg width="24" height="24" viewbox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M19 12H5M12 19l-7-7 7-7" />
      </svg>
     </div>
     <div class="header-title">
      Panel Admin
     </div>
     <div style="width: 24px;"></div>
    </div>
    <div class="content-wrapper">
     <div style="text-align: center; margin-bottom: 25px;">
      <div style="width: 70px; height: 70px; background: linear-gradient(135deg, #f5c857, #f093fb); border-radius: 16px; margin: 0 auto 15px; display: flex; align-items: center; justify-content: center; box-shadow: 0 8px 20px rgba(240, 147, 251, 0.4);">
       <svg width="35" height="35" viewbox="0 0 24 24" fill="none" stroke="white" stroke-width="2"><path d="M12 3c1.657 0 3 1.343 3 3s-1.343 3-3 3-3-1.343-3-3 1.343-3 3-3zM12 11c2.209 0 4 1.791 4 4s-1.791 4-4 4-4-1.791-4-4 1.791-4 4-4zM14 17c-1 0-2 1-3 1s-2-1-3-1M9 20h6" />
       </svg>
      </div>
      <h2 class="text-white text-xl font-bold mb-2" id="admin-welcome">Selamat datang Admin 👋</h2>
      <p class="text-white" style="opacity: 0.8; font-size: 14px;">Kelola pengguna dan password aplikasi</p>
     </div>
     <div class="glass-card" style="margin-bottom: 20px;">
      <h3 class="text-white font-bold mb-4" style="font-size: 16px;">Tambah Pengguna Baru</h3>
      <div style="margin-bottom: 15px;"><label for="new-username" class="text-white block mb-2" style="font-size: 14px;">Username</label> <input type="text" class="glass-input" placeholder="Masukkan username baru" id="new-username">
      </div>
      <div style="margin-bottom: 20px;"><label for="new-password" class="text-white block mb-2" style="font-size: 14px;">Password</label> <input type="password" class="glass-input" placeholder="Masukkan password" id="new-password">
      </div><button class="glass-button" onclick="addNewUser()" style="background: linear-gradient(135deg, #06d6a0, #11998e);">
       <svg width="20" height="20" viewbox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="display: inline-block; margin-right: 8px; vertical-align: middle;"><path d="M12 5v14M5 12h14" />
       </svg> Tambah User </button>
      <p class="text-white" style="opacity: 0.6; font-size: 12px; margin-top: 10px; font-style: italic;">Username minimal 3 karakter, Password minimal 4 karakter</p>
     </div>
     <div class="glass-card">
      <h3 class="text-white font-bold mb-4" style="font-size: 16px;">Daftar Pengguna Terdaftar</h3>
      <div id="users-list" style="display: grid; gap: 10px;">
      </div>
     </div>
     <div class="glass-card" style="margin-top: 20px; background: linear-gradient(135deg, rgba(240, 147, 251, 0.2), rgba(245, 87, 108, 0.2));">
      <h3 class="text-white font-bold mb-4" style="font-size: 16px;">⚙️ Pengaturan Admin</h3>
      <p class="text-white" style="opacity: 0.8; font-size: 13px; margin-bottom: 15px;">Kelola kredensial akun admin Anda</p><button class="glass-button" onclick="editAdminCredentials()" style="background: linear-gradient(135deg, #f093fb, #f5576c);">
       <svg width="20" height="20" viewbox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="display: inline-block; margin-right: 8px; vertical-align: middle;"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7" /> <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z" />
       </svg> Edit Username &amp; Password Admin </button>
     </div>
    </div>
   </div><!-- Page 6: Hasil Analisis -->
   <div id="page-6" class="page-container">
    <div class="header-bar">
     <div class="back-button" onclick="showPage('page-4')">
      <svg width="24" height="24" viewbox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M19 12H5M12 19l-7-7 7-7" />
      </svg>
     </div>
     <div class="header-title">
      Hasil Analisis
     </div>
     <div style="width: 24px;"></div>
    </div>
    <div class="content-wrapper">
     <div class="glass-card" style="margin-bottom: 20px; background: linear-gradient(135deg, rgba(102, 126, 234, 0.3), rgba(118, 75, 162, 0.3));">
      <h3 class="text-white font-bold mb-2" id="result-student-name">-</h3>
      <p class="text-white" style="opacity: 0.85; font-size: 14px; margin-bottom: 8px;" id="result-student-school">-</p>
      <p class="text-white" style="opacity: 0.7; font-size: 13px;">Skor TO: <span class="font-bold text-xl" id="result-student-score">-</span></p>
     </div>
     <div id="analysis-results" style="display: grid; gap: 15px;">
     </div>
     <div style="display: grid; gap: 12px; margin-top: 20px;"><button class="glass-button" onclick="printResults()" style="background: linear-gradient(135deg, #667eea, #764ba2);">
       <svg width="20" height="20" viewbox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="display: inline-block; margin-right: 8px; vertical-align: middle;"><polyline points="6 9 6 2 18 2 18 9" /> <path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2" /> <rect x="6" y="14" width="12" height="8" />
       </svg> Print Hasil </button> <button class="glass-button" onclick="shareResults()" style="background: linear-gradient(135deg, #f093fb, #f5576c);">
       <svg width="20" height="20" viewbox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="display: inline-block; margin-right: 8px; vertical-align: middle;"><circle cx="18" cy="5" r="3" /> <circle cx="6" cy="12" r="3" /> <circle cx="18" cy="19" r="3" /> <line x1="8.59" y1="13.51" x2="15.42" y2="17.49" /> <line x1="15.41" y1="6.51" x2="8.59" y2="10.49" />
       </svg> Share Hasil </button> <button class="glass-button" onclick="startNewConsultation()" style="background: linear-gradient(135deg, #06d6a0, #11998e);">
       <svg width="20" height="20" viewbox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="display: inline-block; margin-right: 8px; vertical-align: middle;"><path d="M12 5v14M5 12h14" />
       </svg> Konsultasi Baru </button>
     </div>
    </div>
   </div><!-- Bottom Navigation -->
   <div class="bottom-nav" id="bottom-nav" style="display: none;">
    <div class="nav-item" onclick="showPage('page-1')" data-page="page-1">
     <svg width="24" height="24" viewbox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z" />
     </svg><span>Home</span>
    </div>
    <div class="nav-item" onclick="showPage('page-2')" data-page="page-2">
     <svg width="24" height="24" viewbox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /> <polyline points="17 8 12 3 7 8" /> <line x1="12" y1="3" x2="12" y2="15" />
     </svg><span>Upload</span>
    </div>
    <div class="nav-item" onclick="showPage('page-3')" data-page="page-3">
     <svg width="24" height="24" viewbox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2" /> <circle cx="12" cy="7" r="4" />
     </svg><span>Siswa</span>
    </div>
    <div class="nav-item" onclick="showPage('page-4')" data-page="page-4">
     <svg width="24" height="24" viewbox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2" ry="2" /> <line x1="16" y1="2" x2="16" y2="6" /> <line x1="8" y1="2" x2="8" y2="6" /> <line x1="3" y1="10" x2="21" y2="10" />
     </svg><span>Pilih</span>
    </div>
    <div class="nav-item" onclick="showPage('page-6')" data-page="page-6">
     <svg width="24" height="24" viewbox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12" />
     </svg><span>Hasil</span>
    </div>
    <div class="nav-item" id="logout-nav-item" style="cursor: pointer;">
     <svg width="24" height="24" viewbox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4" /> <polyline points="16 17 21 12 16 7" /> <line x1="21" y1="12" x2="9" y2="12" />
     </svg><span>Logout</span>
    </div>
   </div><!-- Logout Confirmation Modal -->
   <div id="logout-modal" class="modal-overlay" style="display: none;">
    <div class="modal-content">
     <div class="glass-card">
      <div style="text-align: center; margin-bottom: 20px;">
       <div style="width: 70px; height: 70px; background: linear-gradient(135deg, #dc2626, #991b1b); border-radius: 50%; margin: 0 auto 15px; display: flex; align-items: center; justify-content: center; box-shadow: 0 8px 20px rgba(220, 38, 38, 0.4);">
        <svg width="35" height="35" viewbox="0 0 24 24" fill="none" stroke="white" stroke-width="2"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4" /> <polyline points="16 17 21 12 16 7" /> <line x1="21" y1="12" x2="9" y2="12" />
        </svg>
       </div>
       <h2 class="text-white text-xl font-bold mb-2">Konfirmasi Logout</h2>
       <p class="text-white" style="opacity: 0.8; font-size: 14px;">Apakah Anda yakin ingin keluar dari aplikasi?</p>
      </div>
      <div style="display: grid; gap: 12px;"><button class="glass-button" onclick="confirmLogout()" style="background: linear-gradient(135deg, #dc2626, #991b1b);">
        <svg width="20" height="20" viewbox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="display: inline-block; margin-right: 8px; vertical-align: middle;"><polyline points="20 6 9 17 4 12" />
        </svg> Ya, Logout </button> <button class="glass-button" onclick="hideLogoutModal()" style="background: linear-gradient(135deg, #4b5563, #374151);"> Batal </button>
      </div>
     </div>
    </div>
   </div><!-- Passing Grade Modal -->
   <div id="pg-modal" class="pg-modal-overlay">
    <div class="pg-modal-content">
     <div class="pg-modal-header">
      <div class="pg-modal-title" id="pg-modal-title">
       Passing Grade
      </div>
      <div class="pg-modal-close" onclick="closePGModal()">
       <svg width="24" height="24" viewbox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18" /> <line x1="6" y1="6" x2="18" y2="18" />
       </svg>
      </div>
     </div>
     <div class="pg-modal-body" id="pg-list">
     </div>
    </div>
   </div>
  </div>
  <script>
        const APP_VERSION = '1.4.9';
        const SESSION_KEY = 'snbt_session';
        const LOGIN_KEY = 'snbt_login';
        const DATA_KEY = 'snbt_master_data';
        
        const Storage = {
            set: function(key, value) {
                try {
                    const data = JSON.stringify(value);
                    localStorage.setItem(key, data);
                    sessionStorage.setItem(key, data);
                    return true;
                } catch (e) {
                    return false;
                }
            },
            
            get: function(key) {
                try {
                    let data = localStorage.getItem(key);
                    if (!data) {
                        data = sessionStorage.getItem(key);
                    }
                    return data ? JSON.parse(data) : null;
                } catch (e) {
                    return null;
                }
            },
            
            remove: function(key) {
                try {
                    localStorage.removeItem(key);
                    sessionStorage.removeItem(key);
                    return true;
                } catch (e) {
                    return false;
                }
            },
            
            clear: function() {
                try {
                    localStorage.clear();
                    sessionStorage.clear();
                    return true;
                } catch (e) {
                    return false;
                }
            }
        };

        const Cookie = {
            set: function(name, value, days) {
                try {
                    days = days || 365;
                    const expires = new Date();
                    expires.setTime(expires.getTime() + (days * 24 * 60 * 60 * 1000));
                    const cookieValue = encodeURIComponent(JSON.stringify(value));
                    document.cookie = name + '=' + cookieValue + ';expires=' + expires.toUTCString() + ';path=/;SameSite=Lax';
                    return true;
                } catch (e) {
                    return false;
                }
            },
            
            get: function(name) {
                try {
                    const nameEQ = name + '=';
                    const ca = document.cookie.split(';');
                    for(let i = 0; i < ca.length; i++) {
                        let c = ca[i];
                        while (c.charAt(0) === ' ') c = c.substring(1, c.length);
                        if (c.indexOf(nameEQ) === 0) {
                            const value = decodeURIComponent(c.substring(nameEQ.length, c.length));
                            return JSON.parse(value);
                        }
                    }
                    return null;
                } catch (e) {
                    return null;
                }
            },
            
            remove: function(name) {
                try {
                    document.cookie = name + '=;expires=Thu, 01 Jan 1970 00:00:00 UTC;path=/;';
                    return true;
                } catch (e) {
                    return false;
                }
            }
        };

        function showLoading() {
            const loading = document.getElementById('loading-indicator');
            if (loading) loading.style.display = 'block';
        }

        function hideLoading() {
            const loading = document.getElementById('loading-indicator');
            if (loading) loading.style.display = 'none';
        }

        let masterData = [];
        let studentScore = 0;
        let studentName = '';
        let studentSchool = '';
        let currentChoice = 0;
        let currentFilteredData = [];
        let baseFilteredData = [];
        let selectedChoices = {
            choice1: null,
            choice2: null,
            choice3: null,
            choice4: null
        };
        let isLoggedIn = false;
        let isAdmin = false;
        let userSessionId = null;
        let registeredUsers = {};

        let ADMIN_USERNAME = 'attin';
        let ADMIN_PASSWORD = 'snbt2026';

        function generateSessionId() {
            return 'session_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
        }

        // Backend API Configuration
        const API_BASE_URL = 'https://your-api.com/api'; // Ganti dengan URL API Anda
        const API_KEY = 'your-api-key-here'; // Ganti dengan API key Anda
        const USE_LOCAL_API = true; // Set false untuk menggunakan API eksternal
        
        let dataInitialized = false;
        let allDataRecords = [];
        let isDataSdkReady = false;
        
        async function initializeDataSDK() {
            if (dataInitialized) return;
            dataInitialized = true;

            // Gunakan localStorage sebagai fallback utama
            loadRegisteredUsersFromStorage();
            
            // Jika USE_LOCAL_API = false, sync dengan backend
            if (!USE_LOCAL_API) {
                try {
                    await syncUsersFromBackend();
                } catch (e) {
                    console.error('Failed to sync from backend, using local storage:', e);
                }
            }
        }

        function loadRegisteredUsersFromStorage() {
            const stored = Storage.get('registered_users');
            if (stored && typeof stored === 'object') {
                registeredUsers = stored;
            } else {
                registeredUsers = {};
            }
        }

        async function syncUsersFromBackend() {
            try {
                const response = await fetch(API_BASE_URL + '/users', {
                    method: 'GET',
                    headers: {
                        'Authorization': 'Bearer ' + API_KEY,
                        'Content-Type': 'application/json'
                    }
                });

                if (!response.ok) {
                    throw new Error('Failed to fetch users from backend: ' + response.status);
                }

                const data = await response.json();
                
                if (data.success && Array.isArray(data.users)) {
                    const newRegisteredUsers = {};
                    data.users.forEach(function(user) {
                        if (user.username && user.password) {
                            newRegisteredUsers[user.username] = user.password;
                        }
                    });
                    
                    registeredUsers = newRegisteredUsers;
                    Storage.set('registered_users', registeredUsers);
                    
                    if (isAdmin && document.getElementById('page-admin') && document.getElementById('page-admin').classList.contains('active')) {
                        loadUsersList();
                    }
                }
            } catch (error) {
                console.error('Error syncing users from backend:', error);
                throw error;
            }
        }

        async function saveRegisteredUsers() {
            if (USE_LOCAL_API) {
                // Hanya simpan ke localStorage jika menggunakan local API
                Storage.set('registered_users', registeredUsers);
                return;
            }

            try {
                // Simpan ke backend API
                const response = await fetch(API_BASE_URL + '/users/sync', {
                    method: 'POST',
                    headers: {
                        'Authorization': 'Bearer ' + API_KEY,
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({
                        users: registeredUsers
                    })
                });

                if (!response.ok) {
                    throw new Error('Failed to save users to backend: ' + response.status);
                }

                const result = await response.json();
                
                if (!result.success) {
                    throw new Error(result.message || 'Failed to save users');
                }

                // Simpan juga ke localStorage sebagai backup
                Storage.set('registered_users', registeredUsers);
                
                console.log('Users saved to backend successfully');
            } catch (error) {
                console.error('Error saving users to backend:', error);
                // Fallback: simpan ke localStorage
                Storage.set('registered_users', registeredUsers);
                showToast('Data disimpan secara lokal (backend tidak tersedia)', 'info');
            }
        }

        async function saveUserToBackend(username, password, action) {
            if (USE_LOCAL_API) {
                // Jika menggunakan local API, hanya update registeredUsers
                return true;
            }

            try {
                let endpoint = '';
                let method = '';
                let body = {};

                if (action === 'add') {
                    endpoint = '/users';
                    method = 'POST';
                    body = { username: username, password: password };
                } else if (action === 'update') {
                    endpoint = '/users/' + encodeURIComponent(username);
                    method = 'PUT';
                    body = { password: password };
                } else if (action === 'delete') {
                    endpoint = '/users/' + encodeURIComponent(username);
                    method = 'DELETE';
                }

                const response = await fetch(API_BASE_URL + endpoint, {
                    method: method,
                    headers: {
                        'Authorization': 'Bearer ' + API_KEY,
                        'Content-Type': 'application/json'
                    },
                    body: method !== 'DELETE' ? JSON.stringify(body) : undefined
                });

                if (!response.ok) {
                    throw new Error('Backend error: ' + response.status);
                }

                const result = await response.json();
                return result.success === true;
            } catch (error) {
                console.error('Error with backend operation:', error);
                return false;
            }
        }

        function initApp() {
            try {
                // Initialize Data SDK first for persistent user storage
                if (window.dataSdk) {
                    initializeDataSDK().then(function() {
                        // After SDK initialized, wait a moment and verify data is loaded
                        setTimeout(function() {
                            if (Object.keys(registeredUsers).length === 0) {
                                loadRegisteredUsersFromStorage();
                            }
                        }, 500);
                    }).catch(function() {
                        // If SDK fails, fallback to localStorage
                        loadRegisteredUsersFromStorage();
                    });
                } else {
                    // SDK not available, use localStorage
                    loadRegisteredUsersFromStorage();
                }
                
                // Load admin credentials if saved
                const savedAdminCreds = Storage.get('admin_credentials');
                if (savedAdminCreds) {
                    ADMIN_USERNAME = savedAdminCreds.username;
                    ADMIN_PASSWORD = savedAdminCreds.password;
                }
                
                const storageLogin = Storage.get(LOGIN_KEY);
                const sessionLogin = sessionStorage.getItem(LOGIN_KEY);
                
                if (storageLogin || sessionLogin) {
                    isLoggedIn = true;
                    const loginData = storageLogin || JSON.parse(sessionLogin);
                    isAdmin = loginData.isAdmin || false;
                    
                    const savedData = Storage.get(DATA_KEY);
                    if (savedData && Array.isArray(savedData)) {
                        masterData = savedData;
                    }
                    
                    // Show bottom nav if logged in
                    const bottomNav = document.getElementById('bottom-nav');
                    if (bottomNav) {
                        bottomNav.style.display = 'flex';
                    }
                } else {
                    isLoggedIn = false;
                    isAdmin = false;
                    const bottomNav = document.getElementById('bottom-nav');
                    if (bottomNav) {
                        bottomNav.style.display = 'none';
                    }
                }
            } catch (e) {
                console.error('Init error:', e);
                isLoggedIn = false;
                isAdmin = false;
                loadRegisteredUsersFromStorage();
            }
        }

        function formatPG(value) {
            if (value === null || value === undefined || isNaN(value)) return '0';
            const rounded = Math.round(value * 1000) / 1000;
            if (Number.isInteger(rounded)) return rounded.toString();
            let formatted = rounded.toFixed(3);
            formatted = formatted.replace(/(\.\d{2})\d*?0+$/, '$1');
            return formatted;
        }

        /**
         * Rumus Perhitungan Peluang Diterima
         * 
         * Komponen Penilaian:
         * 1. Skor Akademik (60%):
         *    - Membandingkan skor TO siswa dengan Passing Grade
         *    - Semakin tinggi skor TO vs PG, semakin besar peluang
         *    - Rumus: (Skor TO / PG) × 100% × 0.6
         * 
         * 2. Daya Saing (40%):
         *    - Membandingkan kapasitas dengan jumlah peminat
         *    - Semakin besar kapasitas vs peminat, semakin besar peluang
         *    - Rumus: (Kapasitas / Peminat) × 100% × 0.4
         * 
         * Total Peluang = Skor Akademik + Daya Saing
         * 
         * Kategori Peluang:
         * - Sangat Tinggi (>80%): Warna Hijau (#06d6a0)
         * - Tinggi (60-80%): Warna Kuning (#f5c857)
         * - Sedang (40-60%): Warna Oranye (#f97316)
         * - Rendah (20-40%): Warna Merah Muda (#f5576c)
         * - Sangat Rendah (<20%): Warna Merah (#dc2626)
         */
        function calculatePassChance(studentScore, pg, capacity, applicants) {
            // Validasi input
            if (!studentScore || !pg || !capacity || !applicants) {
                return null; // Program baru atau data tidak lengkap
            }

            // Pastikan nilai lebih dari 0
            if (studentScore <= 0 || pg <= 0 || capacity <= 0 || applicants <= 0) {
                return null;
            }

            // 1. Komponen Skor Akademik (60%)
            // Rumus: (Skor TO / PG) × 100% × 0.6
            // Dibatasi max 100% untuk komponen ini
            const academicScore = Math.min((studentScore / pg) * 100, 100);
            const academicComponent = academicScore * 0.6;

            // 2. Komponen Daya Saing (40%)
            // Rumus: (Kapasitas / Peminat) × 100% × 0.4
            // Dibatasi max 100% untuk komponen ini
            const competitionRatio = Math.min((capacity / applicants) * 100, 100);
            const competitionComponent = competitionRatio * 0.4;

            // Total peluang diterima
            const totalChance = academicComponent + competitionComponent;
            
            // Pastikan hasil dalam range 0-100%
            return Math.min(Math.max(totalChance, 0), 100);
        }

        /**
         * Fungsi untuk menentukan warna berdasarkan persentase peluang
         */
        function getChanceColor(chance) {
            if (chance >= 80) return '#06d6a0'; // Hijau - Sangat Tinggi
            if (chance >= 60) return '#f5c857'; // Kuning - Tinggi
            if (chance >= 40) return '#f97316'; // Oranye - Sedang
            if (chance >= 20) return '#f5576c'; // Merah Muda - Rendah
            return '#dc2626'; // Merah - Sangat Rendah
        }

        /**
         * Fungsi untuk menentukan kategori peluang berdasarkan persentase
         */
        function getChanceCategory(chance) {
            if (chance >= 80) return 'Sangat Tinggi';
            if (chance >= 60) return 'Tinggi';
            if (chance >= 40) return 'Sedang';
            if (chance >= 20) return 'Rendah';
            return 'Sangat Rendah';
        }

        function showToast(message, type) {
            type = type || 'info';
            const toast = document.createElement('div');
            toast.className = 'toast';
            toast.textContent = message;
            if (type === 'success') {
                toast.style.background = 'linear-gradient(135deg, #06d6a0, #11998e)';
            } else if (type === 'error') {
                toast.style.background = 'linear-gradient(135deg, #f5576c, #f093fb)';
            }
            document.body.appendChild(toast);
            setTimeout(function() { 
                if (toast && toast.parentNode) toast.remove(); 
            }, 3000);
        }

        function showPage(pageId) {
            if (!isLoggedIn && pageId !== 'page-1') {
                showToast('Silakan login terlebih dahulu', 'error');
                return;
            }

            document.querySelectorAll('.page-container').forEach(function(page) {
                page.classList.remove('active');
            });
            const targetPage = document.getElementById(pageId);
            if (targetPage) targetPage.classList.add('active');

            document.querySelectorAll('.nav-item').forEach(function(item) {
                item.classList.remove('active');
                if (item.dataset.page === pageId) {
                    item.classList.add('active');
                }
            });

            const bottomNav = document.getElementById('bottom-nav');
            if (bottomNav) {
                if (pageId === 'page-1') {
                    bottomNav.style.display = 'none';
                } else {
                    bottomNav.style.display = 'flex';
                }
            }

            window.scrollTo(0, 0);
        }

        function handleLogin(event) {
            if (event) {
                event.preventDefault();
                event.stopPropagation();
            }
            
            const usernameInput = document.getElementById('username-input');
            const passwordInput = document.getElementById('password-input');
            
            if (!usernameInput || !passwordInput) {
                showToast('Error: Form tidak ditemukan', 'error');
                return false;
            }
            
            const username = usernameInput.value.trim();
            const password = passwordInput.value.trim();
            
            if (username === '') {
                showToast('Masukkan username', 'error');
                usernameInput.focus();
                return false;
            }
            
            if (password === '') {
                showToast('Masukkan password', 'error');
                passwordInput.focus();
                return false;
            }
            
            // Disable button immediately to prevent double-click
            const loginBtn = event.target.querySelector('button[type="submit"]') || document.querySelector('.glass-button');
            if (loginBtn) {
                loginBtn.disabled = true;
                loginBtn.style.opacity = '0.6';
            }
            
            let isValidLogin = false;
            let isAdminLogin = false;
            let matchedUsername = username;
            
            // Check admin credentials (case-insensitive username only)
            if (username.toLowerCase() === ADMIN_USERNAME.toLowerCase() && password === ADMIN_PASSWORD) {
                isValidLogin = true;
                isAdminLogin = true;
                matchedUsername = ADMIN_USERNAME;
            } else {
                // Check registered users (case-sensitive for exact match)
                for (const registeredUsername in registeredUsers) {
                    if (registeredUsername === username && registeredUsers[registeredUsername] === password) {
                        isValidLogin = true;
                        isAdminLogin = false;
                        matchedUsername = registeredUsername;
                        break;
                    }
                }
            }
            
            if (isValidLogin) {
                showLoading();
                
                isLoggedIn = true;
                isAdmin = isAdminLogin;
                
                const loginData = {
                    username: matchedUsername,
                    timestamp: Date.now(),
                    version: APP_VERSION,
                    isAdmin: isAdminLogin
                };
                
                try {
                    Storage.set(LOGIN_KEY, loginData);
                    Cookie.set(LOGIN_KEY, loginData, 30);
                } catch (e) {
                    console.error('Storage error:', e);
                }
                window._loginState = loginData;
                
                requestAnimationFrame(function() {
                    setTimeout(function() {
                        hideLoading();
                        
                        const bottomNav = document.getElementById('bottom-nav');
                        if (bottomNav) {
                            bottomNav.style.display = 'flex';
                        }
                        
                        document.querySelectorAll('.page-container').forEach(function(page) {
                            page.classList.remove('active');
                        });
                        
                        let targetPage = 'page-2';
                        if (isAdminLogin) {
                            targetPage = 'page-admin';
                            loadUsersList();
                            const welcomeText = document.getElementById('admin-welcome');
                            if (welcomeText) welcomeText.textContent = 'Selamat datang Admin ' + matchedUsername + ' 👋';
                        }
                        
                        const page = document.getElementById(targetPage);
                        if (page) {
                            page.classList.add('active');
                        }
                        
                        document.querySelectorAll('.nav-item').forEach(function(item) {
                            item.classList.remove('active');
                        });
                        
                        showToast(isAdminLogin ? 'Login Admin Berhasil! ✓' : 'Login Berhasil! ✓', 'success');
                        window.scrollTo(0, 0);
                        
                        if (loginBtn) {
                            loginBtn.disabled = false;
                            loginBtn.style.opacity = '1';
                        }
                    }, 300);
                });
                
            } else {
                showToast('Username atau Password Salah', 'error');
                passwordInput.value = '';
                passwordInput.focus();
                
                if (loginBtn) {
                    loginBtn.disabled = false;
                    loginBtn.style.opacity = '1';
                }
            }
            
            return false;
        }

        function showLogoutModal() {
            const modal = document.getElementById('logout-modal');
            if (modal) modal.style.display = 'flex';
        }

        function hideLogoutModal() {
            const modal = document.getElementById('logout-modal');
            if (modal) modal.style.display = 'none';
        }

        function adminLogout() {
            confirmLogout();
        }

        function confirmLogout() {
            hideLogoutModal();
            showLoading();
            
            isLoggedIn = false;
            isAdmin = false;
            Storage.remove(LOGIN_KEY);
            Cookie.remove(LOGIN_KEY);
            sessionStorage.removeItem(LOGIN_KEY);
            delete window._loginState;
            
            Storage.remove('student_data');
            studentName = '';
            studentSchool = '';
            studentScore = 0;
            
            const usernameInput = document.getElementById('username-input');
            const passwordInput = document.getElementById('password-input');
            if (usernameInput) usernameInput.value = '';
            if (passwordInput) passwordInput.value = '';
            
            selectedChoices = {
                choice1: null,
                choice2: null,
                choice3: null,
                choice4: null
            };
            
            hideLoading();
            showToast('Logout berhasil', 'success');
            
            setTimeout(function() {
                showPage('page-1');
            }, 300);
        }

        document.addEventListener('DOMContentLoaded', function() {
            initApp();
            
            const loginForm = document.getElementById('login-form');
            if (loginForm) {
                loginForm.addEventListener('submit', handleLogin);
            }

            // Add logout button listener
            const logoutButton = document.getElementById('logout-nav-item');
            if (logoutButton) {
                logoutButton.addEventListener('click', function(e) {
                    if (e) {
                        e.preventDefault();
                        e.stopPropagation();
                    }
                    showLogoutModal();
                });
            }

            const fileUpload = document.getElementById('file-upload');
            if (fileUpload) {
                fileUpload.addEventListener('change', function(e) {
                    const file = e.target.files[0];
                    if (!file) return;

                    showLoading();
                    document.getElementById('file-name-display').textContent = file.name;

                    const reader = new FileReader();
                    reader.onload = function(event) {
                        try {
                            const data = new Uint8Array(event.target.result);
                            const workbook = XLSX.read(data, { type: 'array' });
                            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                            const jsonData = XLSX.utils.sheet_to_json(firstSheet);

                            if (jsonData.length === 0) {
                                hideLoading();
                                showToast('File kosong atau format tidak sesuai', 'error');
                                return;
                            }

                            let streamColumnName = null;
                            if (jsonData.length > 0) {
                                const firstRow = jsonData[0];
                                const possibleStreamColumns = [
                                    'IPA/IPS', 'ipa/ips', 'IPA / IPS',
                                    'STREAM', 'Stream', 'stream',
                                    'JURUSAN', 'Jurusan', 'jurusan',
                                    'J U R U S A N', 'J u r u s a n'
                                ];
                                
                                for (let i = 0; i < possibleStreamColumns.length; i++) {
                                    const colName = possibleStreamColumns[i];
                                    if (firstRow.hasOwnProperty(colName)) {
                                        const value = String(firstRow[colName] || '').trim().toUpperCase();
                                        if (value.includes('IPA') || value.includes('IPS')) {
                                            streamColumnName = colName;
                                            break;
                                        }
                                    }
                                }
                                
                                if (!streamColumnName) {
                                    for (const key in firstRow) {
                                        const value = String(firstRow[key] || '').trim().toUpperCase();
                                        if (value === 'IPA' || value === 'IPS' || 
                                            value.includes('IPA') || value.includes('IPS')) {
                                            streamColumnName = key;
                                            break;
                                        }
                                    }
                                }
                            }

                            masterData = jsonData.map(function(row) {
                                const capacity = parseInt(row['Daya Tampung Sekarang'] || row['daya tampung sekarang'] || row['Daya Tampung'] || 0);
                                const applicants = parseInt(row['Peminat Sebelumnya'] || row['peminat sebelumnya'] || row.Peminat || 0);
                                const competition = parseFloat(row['DAYA SAING'] || row['Daya Saing'] || 0);
                                let rawPG = row.PASSINGRADE || row['PASSING GRADE'] || row.PassingGrade || row.pg || 0;
                                const pg = Number(rawPG);
                                
                                let streamValue = '';
                                if (streamColumnName) {
                                    streamValue = row[streamColumnName] || '';
                                } else {
                                    streamValue = row['IPA/IPS'] || row.STREAM || row.Stream || row.JURUSAN || row.Jurusan || '';
                                }
                                
                                return {
                                    code: row.KODE || row.kode || '',
                                    majorCategory: row['J U R U S A N'] || row.JURUSAN || row.Jurusan || '',
                                    university: row.UNIVERSITAS || row.Universitas || '',
                                    prodi: row.PRODI || row.Prodi || '',
                                    level: row.JENJANG || row.Jenjang || '',
                                    stream: streamValue,
                                    streamColumnSource: streamColumnName,
                                    capacity: isNaN(capacity) ? 0 : capacity,
                                    applicants: isNaN(applicants) ? 0 : applicants,
                                    competition: isNaN(competition) ? 0 : competition,
                                    pg: isNaN(pg) ? 0 : pg
                                };
                            });

                            Storage.set(DATA_KEY, masterData);
                            Cookie.set(DATA_KEY, masterData, 30);

                            displayPreview(masterData);
                            hideLoading();
                            showToast('Data berhasil diupload!', 'success');

                        } catch (error) {
                            hideLoading();
                            showToast('Error membaca file', 'error');
                        }
                    };
                    reader.readAsArrayBuffer(file);
                });
            }
        }, { passive: true });

        function displayPreview(data) {
            const tbody = document.getElementById('preview-tbody');
            const previewCard = document.getElementById('preview-card');

            const totalRows = data.length;
            const uniqueUniversities = [];
            const uniqueMajors = [];
            const uniqueLevels = [];
            
            data.forEach(function(d) {
                if (d.university && d.university.trim() && uniqueUniversities.indexOf(d.university) === -1) {
                    uniqueUniversities.push(d.university);
                }
                if (d.majorCategory && d.majorCategory.trim() && uniqueMajors.indexOf(d.majorCategory) === -1) {
                    uniqueMajors.push(d.majorCategory);
                }
                if (d.level && d.level.trim() && uniqueLevels.indexOf(d.level) === -1) {
                    uniqueLevels.push(d.level);
                }
            });

            document.getElementById('stat-rows').textContent = totalRows;
            document.getElementById('stat-universities').textContent = uniqueUniversities.length;
            document.getElementById('stat-majors').textContent = uniqueMajors.length;
            document.getElementById('stat-levels').textContent = uniqueLevels.length;

            tbody.innerHTML = '';
            
            const sortedData = data.slice().sort(function(a, b) { return b.pg - a.pg; });
            const previewLimit = Math.min(sortedData.length, 5);

            for (let i = 0; i < previewLimit; i++) {
                const item = sortedData[i];
                const row = document.createElement('tr');
                row.innerHTML = '<td>' + item.university + '</td><td>' + item.prodi + '</td><td><strong>' + formatPG(item.pg) + '</strong></td>';
                tbody.appendChild(row);
            }

            previewCard.style.display = 'block';
        }

        function processStudentData() {
            const name = document.getElementById('student-name').value;
            const school = document.getElementById('student-school').value;
            const score = document.getElementById('student-score').value;
            
            if (!name || !school || !score) {
                showToast('Silakan lengkapi semua data siswa', 'error');
                return;
            }
            
            studentName = name;
            studentSchool = school;
            studentScore = parseFloat(score);
            
            const displayScoreElement = document.getElementById('display-score');
            if (displayScoreElement) {
                displayScoreElement.textContent = studentScore;
            }
            
            selectedChoices = {
                choice1: null,
                choice2: null,
                choice3: null,
                choice4: null
            };
            
            Storage.set('student_data', {
                name: studentName,
                school: studentSchool,
                score: studentScore
            });
            
            showPage('page-4');
        }

        function selectChoice(choice) {
            if (masterData.length === 0) {
                showToast('Silakan upload master data terlebih dahulu', 'error');
                return;
            }

            currentChoice = choice;
            
            let filteredData = [];
            let thresholdScore = studentScore;
            const hasScore = studentScore && studentScore > 0;
            
            if (choice === 1) {
                // Pilihan 1: Jika ada skor → S1 BEBAS | Jika kosong → Semua (S1/D4/D3)
                if (hasScore) {
                    filteredData = masterData.filter(function(d) {
                        return d.level === 'S1';
                    });
                } else {
                    // Tidak ada skor → tampilkan semua data tanpa filter level
                    filteredData = masterData.slice();
                }
            } else if (choice === 2) {
                // Pilihan 2: Butuh skor! Jika kosong tidak bisa lanjut
                if (!hasScore) {
                    showToast('Silakan input skor Try Out untuk Pilihan 2', 'error');
                    return;
                }
                // S1 & D4 • PG ≤ Skor TO
                thresholdScore = studentScore;
                filteredData = masterData.filter(function(d) {
                    if (d.level !== 'S1' && d.level !== 'D4') return false;
                    return d.pg <= thresholdScore;
                });
            } else if (choice === 3) {
                // Pilihan 3: Butuh skor! Jika kosong tidak bisa lanjut
                if (!hasScore) {
                    showToast('Silakan input skor Try Out untuk Pilihan 3', 'error');
                    return;
                }
                // D4 & D3 • PG ≤ (Skor TO - 50)
                // Artinya: skor pilihan 3 harus 50 poin di bawah skor pilihan 2
                thresholdScore = studentScore - 50;
                filteredData = masterData.filter(function(d) {
                    if (d.level !== 'D4' && d.level !== 'D3') return false;
                    return d.pg <= thresholdScore;
                });
            } else if (choice === 4) {
                // Pilihan 4: Butuh skor! Jika kosong tidak bisa lanjut
                if (!hasScore) {
                    showToast('Silakan input skor Try Out untuk Pilihan 4', 'error');
                    return;
                }
                // D3 • PG ≤ (Skor TO - 100)
                // Artinya: skor pilihan 4 harus 100 poin di bawah skor pilihan 2 (50 poin di bawah pilihan 3)
                thresholdScore = studentScore - 100;
                filteredData = masterData.filter(function(d) {
                    if (d.level !== 'D3') return false;
                    return d.pg <= thresholdScore;
                });
            }

            // Hapus program yang sudah dipilih di pilihan sebelumnya (pilihan yang lebih tinggi prioritas)
            let alreadySelectedPrograms = [];
            for (let i = 1; i < choice; i++) {
                const choiceKey = 'choice' + i;
                if (selectedChoices[choiceKey] && Array.isArray(selectedChoices[choiceKey])) {
                    alreadySelectedPrograms = alreadySelectedPrograms.concat(selectedChoices[choiceKey]);
                }
            }

            if (alreadySelectedPrograms.length > 0) {
                filteredData = filteredData.filter(function(program) {
                    return !alreadySelectedPrograms.some(function(selected) {
                        return selected.code === program.code && 
                               selected.university === program.university && 
                               selected.prodi === program.prodi;
                    });
                });
            }

            if (filteredData.length === 0) {
                showToast('Tidak ada program studi yang memenuhi kriteria', 'error');
                return;
            }

            filteredData.sort(function(a, b) { return b.pg - a.pg; });
            baseFilteredData = filteredData;
            currentFilteredData = filteredData;

            populateFilters(filteredData);
            
            document.getElementById('filter-university').value = '';
            document.getElementById('filter-level').value = '';
            document.getElementById('filter-stream').value = '';
            document.getElementById('search-prodi').value = '';
            
            displayPrograms(filteredData);
            
            showPage('page-5');
        }

        function populateFilters(data) {
            const universities = [];
            const levels = [];
            
            data.forEach(function(d) {
                if (universities.indexOf(d.university) === -1) universities.push(d.university);
                if (levels.indexOf(d.level) === -1) levels.push(d.level);
            });
            
            universities.sort();
            levels.sort();

            const universitySelect = document.getElementById('filter-university');
            universitySelect.innerHTML = '<option value="">Semua Universitas</option>';
            universities.forEach(function(uni) {
                const option = document.createElement('option');
                option.value = uni;
                option.textContent = uni;
                universitySelect.appendChild(option);
            });

            const levelSelect = document.getElementById('filter-level');
            levelSelect.innerHTML = '<option value="">Semua Jenjang</option>';
            levels.forEach(function(level) {
                const option = document.createElement('option');
                option.value = level;
                option.textContent = level;
                levelSelect.appendChild(option);
            });
        }

        let lastFilterNotification = 0;
        let lastSearchTerm = '';
        
        function applyFilters() {
            const searchInput = document.getElementById('search-prodi');
            const streamSelect = document.getElementById('filter-stream');
            const universitySelect = document.getElementById('filter-university');
            const levelSelect = document.getElementById('filter-level');
            
            if (!searchInput || !streamSelect || !universitySelect || !levelSelect) {
                return;
            }
            
            const searchTerm = searchInput.value.toLowerCase().trim();
            const streamFilter = streamSelect.value.trim();
            const universityFilter = universitySelect.value.trim();
            const levelFilter = levelSelect.value.trim();

            let filtered = baseFilteredData.slice();

            if (searchTerm && searchTerm.length > 0) {
                filtered = filtered.filter(function(d) {
                    const majorCategoryName = (d.majorCategory || '').toLowerCase();
                    const prodiName = (d.prodi || '').toLowerCase();
                    
                    const searchWords = searchTerm.split(/\s+/).filter(function(w) { return w.length > 0; });
                    
                    return searchWords.every(function(searchWord) {
                        return majorCategoryName.indexOf(searchWord) !== -1 || prodiName.indexOf(searchWord) !== -1;
                    });
                });
            }

            if (streamFilter && streamFilter.length > 0) {
                filtered = filtered.filter(function(d) {
                    const rawStream = d.stream || '';
                    const streamValue = String(rawStream).trim().toUpperCase();
                    const filterValue = streamFilter.toUpperCase();
                    
                    if (!streamValue || streamValue === '') {
                        return false;
                    }
                    
                    const exactMatch = streamValue === filterValue;
                    const containsMatch = streamValue.indexOf(filterValue) !== -1;
                    const startsWithMatch = streamValue.indexOf(filterValue) === 0;
                    const endsWithMatch = streamValue.lastIndexOf(filterValue) === streamValue.length - filterValue.length;
                    
                    const streamWords = streamValue.split(/[\s\-_\/\(\)]+/).filter(function(w) { return w.length > 0; });
                    const wordMatch = streamWords.some(function(word) {
                        return word === filterValue || word.indexOf(filterValue) === 0 || filterValue.indexOf(word) === 0;
                    });
                    
                    return exactMatch || containsMatch || startsWithMatch || endsWithMatch || wordMatch;
                });
            }

            if (levelFilter && levelFilter.length > 0 && levelFilter !== '') {
                filtered = filtered.filter(function(d) { return d.level === levelFilter; });
            }

            updateUniversityFilter(filtered);

            if (universityFilter && universityFilter.length > 0 && universityFilter !== '') {
                filtered = filtered.filter(function(d) { return d.university === universityFilter; });
            }

            currentFilteredData = filtered;
            displayPrograms(filtered);
            
            // Show notification only once per search term change
            const now = Date.now();
            if (searchTerm && searchTerm.length > 0 && filtered.length === 0 && searchTerm !== lastSearchTerm && (now - lastFilterNotification) > 500) {
                showToast('Jurusan tidak tersedia', 'error');
                lastFilterNotification = now;
                lastSearchTerm = searchTerm;
            } else if (searchTerm === '') {
                lastSearchTerm = '';
            }
        }

        function updateUniversityFilter(filteredData) {
            const universitySelect = document.getElementById('filter-university');
            if (!universitySelect) return;
            
            const currentValue = universitySelect.value;
            
            const universities = [];
            filteredData.forEach(function(d) {
                if (universities.indexOf(d.university) === -1) {
                    universities.push(d.university);
                }
            });
            
            universities.sort();
            
            universitySelect.innerHTML = '<option value="">Semua Universitas</option>';
            universities.forEach(function(uni) {
                const option = document.createElement('option');
                option.value = uni;
                option.textContent = uni;
                universitySelect.appendChild(option);
            });
            
            // Always preserve the selected university, even if it's not in filtered results
            if (currentValue && currentValue !== '') {
                const optionExists = universities.indexOf(currentValue) !== -1;
                if (optionExists) {
                    universitySelect.value = currentValue;
                } else {
                    // Keep value selected even if option doesn't exist in filtered list
                    // This allows users to keep their selection while searching
                    const missingOption = document.createElement('option');
                    missingOption.value = currentValue;
                    missingOption.textContent = currentValue + ' (tidak ada jurusan)';
                    universitySelect.appendChild(missingOption);
                    universitySelect.value = currentValue;
                }
            } else {
                universitySelect.value = '';
            }
        }

        function updateResultCount() {
            const resultCount = document.getElementById('result-count');
            if (resultCount) {
                const searchInput = document.getElementById('search-prodi');
                const searchTerm = searchInput ? searchInput.value.toLowerCase().trim() : '';
                
                if (currentFilteredData.length === 0 && searchTerm && searchTerm.length > 0) {
                    resultCount.textContent = 'Jurusan tidak tersedia';
                } else {
                    resultCount.textContent = currentFilteredData.length + ' program studi yang disarankan';
                }
            }
        }



        function displayPrograms(programs) {
            const container = document.getElementById('programs-list');
            const resultCount = document.getElementById('result-count');
            
            if (programs.length === 0) {
                container.innerHTML = '<div class="glass-card" style="text-align: center; padding: 40px;"><p class="text-white" style="opacity: 0.6;">Tidak ada program studi yang sesuai</p></div>';
                if (resultCount) resultCount.textContent = '0 program studi yang disarankan';
                return;
            }

            container.innerHTML = '';
            const fragment = document.createDocumentFragment();
            
            for (let i = 0; i < Math.min(programs.length, 100); i++) {
                const program = programs[i];
                const isNewProgram = program.applicants === 0 || program.applicants === null;
                
                const card = document.createElement('div');
                card.className = 'glass-card program-card-item';
                card.style.padding = '15px';
                card.setAttribute('data-index', i);
                
                let badgesHtml = '<span class="status-badge" style="background: rgba(102, 126, 234, 0.4); color: white;">' + program.level + '</span>';
                if (program.stream) {
                    badgesHtml += '<span class="status-badge" style="background: rgba(220, 38, 38, 0.4); color: white;">' + program.stream + '</span>';
                }
                badgesHtml += '<span class="status-badge" style="background: rgba(6, 214, 160, 0.4); color: white;">PG: ' + formatPG(program.pg) + '</span>';
                if (isNewProgram) {
                    badgesHtml += '<span class="status-badge" style="background: rgba(240, 147, 251, 0.4); color: white;">Baru</span>';
                }
                
                let majorHtml = '';
                if (program.majorCategory) {
                    majorHtml = '<p class="text-white" style="opacity: 0.9; font-size: 13px; margin-bottom: 6px; color: white; font-weight: 600;">' + program.majorCategory + '</p>';
                }
                
                card.innerHTML = '<div style="display: flex; gap: 12px; align-items: flex-start; justify-content: space-between;"><div style="flex: 1;"><h3 class="text-white font-bold" style="font-size: 16px; margin-bottom: 10px; color: white; text-transform: uppercase; letter-spacing: 0.3px; line-height: 1.3;">' + program.university + '</h3>' + majorHtml + '<h4 class="text-white font-bold" style="font-size: 14px; margin-bottom: 8px; color: white; opacity: 0.95;">' + program.prodi + '</h4><div style="display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 8px;">' + badgesHtml + '</div><p class="text-white" style="opacity: 0.7; font-size: 11px; color: white;">Kapasitas: ' + program.capacity + ' / Peminat: ' + (isNewProgram ? 'Baru' : program.applicants) + '</p></div><button class="select-btn" data-index="' + i + '" style="padding: 10px 18px; background: linear-gradient(135deg, #06d6a0, #11998e); border: none; border-radius: 10px; color: white; font-weight: 600; cursor: pointer; font-size: 13px; white-space: nowrap; transition: all 0.2s ease; flex-shrink: 0;">Pilih</button></div>';
                
                fragment.appendChild(card);
            }
            
            container.appendChild(fragment);
            
            // Attach click handlers to select buttons
            container.querySelectorAll('.select-btn').forEach(function(btn) {
                btn.addEventListener('click', function() {
                    const index = parseInt(this.getAttribute('data-index'));
                    toggleProgramSelection(index, this);
                });
            });
            
            updateSelectedCount();
            
            // Update result count
            if (resultCount) {
                const searchInput = document.getElementById('search-prodi');
                const searchTerm = searchInput ? searchInput.value.toLowerCase().trim() : '';
                
                if (programs.length === 0 && searchTerm && searchTerm.length > 0) {
                    resultCount.textContent = 'Jurusan tidak tersedia';
                } else {
                    resultCount.textContent = programs.length + ' program studi yang disarankan';
                }
            }
        }
        
        function toggleProgramSelection(index, button) {
            const program = currentFilteredData[index];
            if (!program) return;
            
            // Check if already selected
            const isSelected = button.classList.contains('selected');
            
            if (isSelected) {
                button.classList.remove('selected');
                button.textContent = 'Pilih';
                button.style.background = 'linear-gradient(135deg, #06d6a0, #11998e)';
            } else {
                button.classList.add('selected');
                button.textContent = '✓ Dipilih';
                button.style.background = 'linear-gradient(135deg, #667eea, #764ba2)';
            }
            
            updateSelectedCount();
        }

        function loadUniversityLogos() {
            const logoImages = document.querySelectorAll('.university-logo');
            let loadedCount = 0;
            
            logoImages.forEach(function(img) {
                const universityName = img.dataset.university;
                if (universityName) {
                    loadedCount++;
                } else {
                    loadedCount++;
                }
            });
        }

        function updateSelectedCount() {
            const selectedButtons = document.querySelectorAll('.select-btn.selected');
            let selected = 0;
            selectedButtons.forEach(function(btn) {
                if (btn.classList.contains('selected')) selected++;
            });
            const countElement = document.getElementById('selected-count');
            if (countElement) countElement.textContent = selected;
        }



        function confirmSelection() {
            const selectedButtons = document.querySelectorAll('.select-btn.selected');
            if (selectedButtons.length === 0) {
                showToast('Silakan pilih minimal 1 program studi', 'error');
                return;
            }

            const selectedPrograms = [];
            selectedButtons.forEach(function(button) {
                const index = parseInt(button.getAttribute('data-index'));
                selectedPrograms.push(currentFilteredData[index]);
            });
            
            selectedChoices['choice' + currentChoice] = selectedPrograms;
            
            const checkIcon = document.getElementById('check-' + currentChoice);
            const card = document.getElementById('choice-card-' + currentChoice);
            if (checkIcon) checkIcon.style.display = 'flex';
            if (card) card.classList.add('selected');
            
            let selectedCount = 0;
            let totalPrograms = 0;
            
            for (const key in selectedChoices) {
                const choice = selectedChoices[key];
                if (choice !== null && (Array.isArray(choice) ? choice.length > 0 : true)) {
                    selectedCount++;
                    totalPrograms += Array.isArray(choice) ? choice.length : 1;
                }
            }
            
            const choicesStatus = document.getElementById('choices-status');
            if (choicesStatus) {
                choicesStatus.textContent = selectedCount + ' pilihan dipilih (' + totalPrograms + ' program)';
            }
            
            if (selectedCount > 0) {
                const analyzeBtn = document.getElementById('analyze-button');
                if (analyzeBtn) analyzeBtn.style.display = 'block';
            }
            
            showToast('Pilihan ' + currentChoice + ' berhasil (' + selectedPrograms.length + ' program)!', 'success');
            showPage('page-4');
        }

        function goToAnalysis() {
            let allSelectedPrograms = [];
            for (const key in selectedChoices) {
                const choice = selectedChoices[key];
                if (choice !== null) {
                    if (Array.isArray(choice)) {
                        allSelectedPrograms = allSelectedPrograms.concat(choice);
                    } else {
                        allSelectedPrograms.push(choice);
                    }
                }
            }
            
            if (allSelectedPrograms.length === 0) {
                showToast('Silakan pilih minimal 1 program studi', 'error');
                return;
            }
            
            displayAnalysis(allSelectedPrograms);
            showPage('page-6');
        }

        function displayAnalysis(programs) {
            document.getElementById('result-student-name').textContent = studentName;
            document.getElementById('result-student-school').textContent = studentSchool;
            document.getElementById('result-student-score').textContent = studentScore;

            const container = document.getElementById('analysis-results');
            container.innerHTML = '';

            const choiceLabels = { 1: 'PILIHAN 1', 2: 'PILIHAN 2', 3: 'PILIHAN 3', 4: 'PILIHAN 4' };
            const choiceColors = {
                1: { bg: 'rgba(79, 172, 254, 0.2)', border: 'rgba(79, 172, 254, 0.5)' },
                2: { bg: 'rgba(118, 75, 162, 0.2)', border: 'rgba(118, 75, 162, 0.5)' },
                3: { bg: 'rgba(67, 97, 238, 0.2)', border: 'rgba(67, 97, 238, 0.5)' },
                4: { bg: 'rgba(6, 214, 160, 0.2)', border: 'rgba(6, 214, 160, 0.5)' }
            };

            const fragment = document.createDocumentFragment();
            
            // Display programs grouped by choice
            for (let choiceNum = 1; choiceNum <= 4; choiceNum++) {
                const choicePrograms = selectedChoices['choice' + choiceNum];
                
                if (!choicePrograms || !Array.isArray(choicePrograms) || choicePrograms.length === 0) {
                    continue;
                }

                const choiceSection = document.createElement('div');
                choiceSection.style.marginBottom = '20px';
                
                // Choice header
                const choiceHeader = document.createElement('div');
                choiceHeader.style.cssText = 'background: linear-gradient(135deg, rgba(102, 126, 234, 0.3), rgba(118, 75, 162, 0.3)); padding: 12px 16px; border-radius: 12px; margin-bottom: 12px; border-left: 4px solid rgba(220, 38, 38, 0.6);';
                choiceHeader.innerHTML = '<h3 class="text-white font-bold" style="font-size: 16px; margin-bottom: 4px;">' + choiceLabels[choiceNum] + '</h3><p class="text-white" style="opacity: 0.8; font-size: 12px;">' + choicePrograms.length + ' program dipilih</p>';
                choiceSection.appendChild(choiceHeader);
                
                // Programs for this choice
                choicePrograms.forEach(function(program) {
                    const isNewProgram = program.applicants === 0 || program.applicants === null;
                    
                    // Gunakan fungsi calculatePassChance untuk menghitung peluang
                    const passChance = isNewProgram ? null : calculatePassChance(studentScore, program.pg, program.capacity, program.applicants);
                    const gradeColor = passChance !== null ? getChanceColor(passChance) : '#f093fb';

                    const card = document.createElement('div');
                    card.className = 'glass-card';
                    card.style.marginBottom = '12px';
                    card.style.borderColor = choiceColors[choiceNum].border;
                    card.style.background = choiceColors[choiceNum].bg;
                    
                    let html = '<div style="margin-bottom: 12px;">';
                    
                    // Header dengan university dan percentage side-by-side
                    html += '<div style="display: flex; gap: 12px; align-items: flex-start; justify-content: space-between; margin-bottom: 12px;">';
                    
                    // University name - flex: 1 agar tidak bisa overlap
                    html += '<div style="flex: 1; min-width: 0;">';
                    html += '<h3 class="text-white font-bold" style="font-size: 16px; line-height: 1.3; text-transform: uppercase; letter-spacing: 0.5px; word-break: break-word; margin: 0;">' + program.university + '</h3>';
                    html += '</div>';
                    
                    // Circular percentage badge - flex-shrink: 0 agar tidak pernah mengecil
                    if (passChance !== null) {
                        const circumference = 2 * Math.PI * 28;
                        const strokeDashoffset = circumference - (passChance / 100) * circumference;
                        html += '<div style="display: flex; align-items: center; justify-content: center; flex-shrink: 0;">';
                        html += '<div style="position: relative; width: 70px; height: 70px; display: flex; align-items: center; justify-content: center;">';
                        html += '<svg width="70" height="70" style="position: absolute; transform: rotate(-90deg); filter: drop-shadow(0 2px 8px rgba(0,0,0,0.3));">';
                        html += '<circle cx="35" cy="35" r="28" fill="none" stroke="rgba(255,255,255,0.15)" stroke-width="5"/>';
                        html += '<circle cx="35" cy="35" r="28" fill="none" stroke="' + gradeColor + '" stroke-width="5" stroke-dasharray="' + circumference + '" stroke-dashoffset="' + strokeDashoffset + '" stroke-linecap="round" style="transition: stroke-dashoffset 0.5s ease;"/>';
                        html += '</svg>';
                        html += '<div style="text-align: center; display: flex; align-items: center; gap: 1px; position: relative; z-index: 1;">';
                        html += '<div class="text-white font-bold" style="font-size: 14px; line-height: 1; color: ' + gradeColor + ';">' + Math.round(passChance) + '</div>';
                        html += '<div class="text-white" style="font-size: 9px; opacity: 0.8; line-height: 1;">%</div>';
                        html += '</div>';
                        html += '</div>';
                        html += '</div>';
                    } else {
                        html += '<div style="display: flex; align-items: center; justify-content: center; flex-shrink: 0;">';
                        html += '<div style="position: relative; width: 70px; height: 70px; display: flex; align-items: center; justify-content: center;">';
                        html += '<svg width="70" height="70" style="position: absolute; filter: drop-shadow(0 2px 8px rgba(0,0,0,0.3));">';
                        html += '<circle cx="35" cy="35" r="28" fill="none" stroke="rgba(255,255,255,0.15)" stroke-width="5"/>';
                        html += '<circle cx="35" cy="35" r="28" fill="none" stroke="rgba(240,147,251,0.8)" stroke-width="5"/>';
                        html += '</svg>';
                        html += '<div style="text-align: center; position: relative; z-index: 1;">';
                        html += '<div class="text-white font-bold" style="font-size: 11px; line-height: 1;">BARU</div>';
                        html += '</div>';
                        html += '</div>';
                        html += '</div>';
                    }
                    
                    html += '</div>';
                    
                    if (program.majorCategory) {
                        html += '<p class="text-white" style="opacity: 0.85; font-size: 13px; margin-bottom: 4px; font-weight: 600;">' + program.majorCategory + '</p>';
                    }
                    html += '<p class="text-white" style="opacity: 0.9; font-size: 13px; margin-bottom: 8px;">' + program.prodi + '</p>';
                    html += '<p class="text-white" style="opacity: 0.75; font-size: 12px; margin-bottom: 10px;">' + program.level;
                    if (program.stream) {
                        html += ' | ' + program.stream;
                    }
                    html += ' | PG: ' + formatPG(program.pg) + '</p>';
                    
                    // Academic Score Bar
                    const academicScore = Math.min((studentScore / program.pg) * 100, 100);
                    html += '<div style="margin-bottom: 10px;">';
                    html += '<div style="display: flex; justify-content: space-between; margin-bottom: 6px;">';
                    html += '<span class="text-white" style="opacity: 0.8; font-size: 12px;">Skor vs PG</span>';
                    html += '<span class="text-white font-bold" style="font-size: 12px;">' + studentScore + ' / ' + formatPG(program.pg) + '</span>';
                    html += '</div>';
                    html += '<div style="height: 6px; background: rgba(255,255,255,0.2); border-radius: 10px; overflow: hidden;">';
                    html += '<div style="width: ' + academicScore + '%; height: 100%; background: ' + gradeColor + ';"></div>';
                    html += '</div></div>';
                    
                    // Competition/Capacity Bar
                    if (!isNewProgram) {
                        const capacityCompetition = Math.min((program.capacity / program.applicants) * 100, 100);
                        html += '<div>';
                        html += '<div style="display: flex; justify-content: space-between; margin-bottom: 6px;">';
                        html += '<span class="text-white" style="opacity: 0.8; font-size: 12px;">Daya Tampung</span>';
                        html += '<span class="text-white font-bold" style="font-size: 12px;">' + program.capacity + ' / ' + program.applicants + '</span>';
                        html += '</div>';
                        html += '<div style="height: 6px; background: rgba(255,255,255,0.2); border-radius: 10px; overflow: hidden;">';
                        html += '<div style="width: ' + capacityCompetition + '%; height: 100%; background: #f5576c;"></div>';
                        html += '</div></div>';
                    } else {
                        html += '<p class="text-white" style="opacity: 0.7; font-size: 12px;">Daya Tampung: ' + program.capacity + ' (Program Baru)</p>';
                    }
                    
                    card.innerHTML = html;
                    choiceSection.appendChild(card);
                });
                
                fragment.appendChild(choiceSection);
            }

            container.appendChild(fragment);
        }

        async function fetchUniversityLogo(universityName) {
            if (logoCache[universityName] !== undefined) {
                return logoCache[universityName];
            }
            
            try {
                const searchQuery = encodeURIComponent(universityName);
                const response = await fetch('https://www.wikidata.org/w/api.php?action=wbsearch&search=' + searchQuery + '&language=en&limit=1&format=json&origin=*', { method: 'GET', mode: 'cors' });
                const data = await response.json();
                
                if (!data.search || data.search.length === 0) {
                    logoCache[universityName] = null;
                    return null;
                }
                
                const entityId = data.search[0].id;
                const entityResponse = await fetch('https://www.wikidata.org/wiki/Special:EntityData/' + entityId + '.json', { method: 'GET', mode: 'cors' });
                const entityData = await entityResponse.json();
                const entity = entityData.entities[entityId];
                
                if (entity.claims && entity.claims.P154 && entity.claims.P154.length > 0) {
                    const logoFilename = entity.claims.P154[0].mainsnak.datavalue.value;
                    const logoUrl = 'https://commons.wikimedia.org/wiki/Special:FilePath/' + encodeURIComponent(logoFilename) + '?width=80';
                    logoCache[universityName] = logoUrl;
                    return logoUrl;
                } else if (entity.claims && entity.claims.P18 && entity.claims.P18.length > 0) {
                    const imageFilename = entity.claims.P18[0].mainsnak.datavalue.value;
                    const imageUrl = 'https://commons.wikimedia.org/wiki/Special:FilePath/' + encodeURIComponent(imageFilename) + '?width=80';
                    logoCache[universityName] = imageUrl;
                    return imageUrl;
                }
                
                logoCache[universityName] = null;
                return null;
            } catch (error) {
                logoCache[universityName] = null;
                return null;
            }
        }

        async function printResults() {
            const printContent = generatePrintContent();
            
            const printWindow = window.open('', '_blank', 'width=800,height=600');
            if (!printWindow) {
                showToast('Pop-up diblokir! Izinkan pop-up untuk print', 'error');
                return;
            }
            
            printWindow.document.write(printContent);
            printWindow.document.close();
            
            printWindow.onload = function() {
                printWindow.focus();
                setTimeout(function() {
                    printWindow.print();
                    showToast('Dokumen siap untuk dicetak!', 'success');
                }, 500);
            };
        }
        
        function getConsultantName() {
            try {
                const loginData = Storage.get(LOGIN_KEY) || window._loginState;
                if (loginData && loginData.username) {
                    return loginData.username;
                }
            } catch (e) {
                console.error('Error getting consultant name:', e);
            }
            return 'Konsultan';
        }
        
        function generatePrintContent() {
            const choice1Programs = selectedChoices.choice1 && Array.isArray(selectedChoices.choice1) ? selectedChoices.choice1 : [];
            const choice2Programs = selectedChoices.choice2 && Array.isArray(selectedChoices.choice2) ? selectedChoices.choice2 : [];
            const choice3Programs = selectedChoices.choice3 && Array.isArray(selectedChoices.choice3) ? selectedChoices.choice3 : [];
            const choice4Programs = selectedChoices.choice4 && Array.isArray(selectedChoices.choice4) ? selectedChoices.choice4 : [];
            
            let html = '<!DOCTYPE html><html><head>';
            html += '<meta charset="UTF-8">';
            html += '<title>Hasil Konsultasi SNBT - ' + studentName + '</title>';
            html += '<style>';
            html += '@page { size: A4; margin: 8mm; }';
            html += '* { margin: 0; padding: 0; box-sizing: border-box; }';
            html += 'body { font-family: "Segoe UI", Arial, sans-serif; background: white; color: #1a1a1a; line-height: 1.3; font-size: 9pt; }';
            html += '.page { width: 210mm; min-height: 297mm; padding: 8mm; background: white; }';
            html += '.header-section { text-align: center; padding: 10px 0; border-bottom: 3px solid #dc2626; margin-bottom: 12px; background: linear-gradient(135deg, #fef2f2 0%, #fee2e2 100%); border-radius: 6px; }';
            html += '.logo-print { width: 50px; height: 50px; margin: 0 auto 8px; }';
            html += '.logo-print img { width: 100%; height: 100%; object-fit: contain; border-radius: 50%; border: 2px solid #dc2626; }';
            html += '.main-title { color: #dc2626; font-size: 18px; font-weight: bold; margin-bottom: 3px; text-transform: uppercase; letter-spacing: 0.5px; }';
            html += '.subtitle { color: #666; font-size: 10px; margin-bottom: 2px; }';
            html += '.version-tag { display: inline-block; background: #dc2626; color: white; padding: 2px 8px; border-radius: 10px; font-size: 8px; margin-top: 4px; }';
            html += '.info-section { background: linear-gradient(135deg, #fef2f2 0%, #fee2e2 100%); border-left: 4px solid #dc2626; padding: 10px 12px; margin-bottom: 12px; border-radius: 5px; }';
            html += '.info-grid { display: grid; grid-template-columns: 1fr 1fr 1fr 1fr; gap: 8px; }';
            html += '.info-item { padding: 3px 0; }';
            html += '.info-label { font-size: 8px; color: #666; text-transform: uppercase; letter-spacing: 0.2px; margin-bottom: 2px; font-weight: 600; }';
            html += '.info-value { font-size: 10px; color: #1a1a1a; font-weight: bold; }';
            html += '.choices-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 10px; }';
            html += '.choice-section { break-inside: avoid; }';
            html += '.choice-title { color: white; font-size: 11px; font-weight: bold; margin-bottom: 8px; padding: 6px 10px; background: linear-gradient(135deg, #dc2626 0%, #991b1b 100%); border-radius: 5px; text-align: center; text-transform: uppercase; letter-spacing: 0.4px; }';
            html += '.program-card { border: 2px solid #dc2626; border-radius: 6px; padding: 8px; margin-bottom: 8px; break-inside: avoid; background: linear-gradient(135deg, #fef2f2 0%, #ffffff 100%); position: relative; overflow: hidden; box-shadow: 0 1px 4px rgba(220, 38, 38, 0.1); }';
            html += '.program-card::before { content: ""; position: absolute; top: 0; left: 0; width: 3px; height: 100%; background: linear-gradient(180deg, #dc2626 0%, #991b1b 100%); }';
            html += '.program-header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 6px; padding-left: 6px; gap: 8px; }';
            html += '.program-header-left { flex: 1; }';
            html += '.peluang-box { display: flex; flex-direction: column; align-items: center; background: white; padding: 4px 8px; border-radius: 5px; border: 2px solid #06d6a0; min-width: 45px; flex-shrink: 0; }';
            html += '.peluang-label { font-size: 6px; color: #666; text-transform: uppercase; margin-bottom: 1px; font-weight: 600; }';
            html += '.peluang-value { font-size: 14px; font-weight: bold; color: #06d6a0; line-height: 1; }';
            html += '.peluang-box.new { border-color: #f093fb; }';
            html += '.peluang-box.new .peluang-value { color: #f093fb; font-size: 9px; }';
            html += '.university-name { font-size: 10px; font-weight: bold; color: #dc2626; margin-bottom: 5px; padding-left: 6px; text-transform: uppercase; line-height: 1.2; }';
            html += '.program-info { padding-left: 6px; margin-bottom: 6px; }';
            html += '.program-value { font-size: 9px; color: #1a1a1a; margin-bottom: 3px; font-weight: 600; line-height: 1.2; }';
            html += '.badges-row { display: flex; gap: 3px; flex-wrap: wrap; margin-bottom: 6px; padding-left: 6px; }';
            html += '.badge { padding: 2px 6px; border-radius: 8px; font-size: 7px; font-weight: bold; }';
            html += '.badge-level { background: #dbeafe; color: #1e40af; }';
            html += '.badge-stream { background: #fef3c7; color: #92400e; }';
            html += '.badge-pg { background: #d1fae5; color: #065f46; }';
            html += '.stats-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 4px; background: white; padding: 6px; border-radius: 5px; margin-left: 6px; border: 1px solid #e5e7eb; }';
            html += '.stat-box { text-align: center; padding: 4px; border-radius: 3px; background: #f9fafb; }';
            html += '.stat-label { font-size: 6px; color: #666; text-transform: uppercase; margin-bottom: 2px; font-weight: 600; letter-spacing: 0.2px; }';
            html += '.stat-value { font-size: 10px; font-weight: bold; color: #dc2626; line-height: 1; }';
            html += '.signature-section { display: grid; grid-template-columns: 1fr 1fr; gap: 40px; margin-top: 25px; margin-bottom: 20px; padding: 0 30px; }';
            html += '.signature-box { text-align: center; }';
            html += '.signature-label { font-size: 12px; color: #1a1a1a; margin-bottom: 15px; font-weight: bold; text-transform: uppercase; }';
            html += '.signature-space { height: 80px; border-bottom: 2px solid #1a1a1a; margin-bottom: 10px; }';
            html += '.signature-name { font-size: 11px; color: #1a1a1a; font-weight: 600; }';
            html += '.footer-section { margin-top: 12px; padding-top: 10px; border-top: 2px solid #0ea5e9; text-align: center; background: linear-gradient(135deg, #e0f2fe 0%, #bae6fd 100%); border-radius: 5px; padding-bottom: 10px; }';
            html += '.footer-logo { width: 35px; height: 35px; margin: 0 auto 6px; }';
            html += '.footer-logo img { width: 100%; height: 100%; object-fit: contain; border-radius: 50%; }';
            html += '.footer-title { font-size: 12px; font-weight: bold; color: #0284c7; margin-bottom: 3px; }';
            html += '.footer-text { font-size: 8px; color: #666; margin-bottom: 2px; }';
            html += '.footer-date { font-size: 7px; color: #999; margin-top: 4px; font-style: italic; }';
            html += '.empty-choice { text-align: center; padding: 15px; color: #999; font-size: 9px; font-style: italic; background: #f9fafb; border-radius: 5px; border: 2px dashed #e5e7eb; }';
            html += '@media print { ';
            html += 'body { print-color-adjust: exact; -webkit-print-color-adjust: exact; }';
            html += '.page { page-break-after: auto; }';
            html += '.program-card { page-break-inside: avoid; }';
            html += '.choice-section { page-break-inside: avoid; }';
            html += '}';
            html += '</style>';
            html += '</head><body>';
            
            html += '<div class="page">';
            
            html += '<div class="header-section">';
            html += '<div class="logo-print"><img src="https://iili.io/KsKbLep.png" alt="Logo Bimbel Attin"></div>';
            html += '<div class="main-title">Hasil Konsultasi SNBT</div>';
            html += '<div class="subtitle">Bimbel Attin - Sistem Rekomendasi Program Studi</div>';
            html += '<div class="version-tag">Version ' + APP_VERSION + '</div>';
            html += '</div>';
            
            html += '<div class="info-section">';
            html += '<div class="info-grid">';
            html += '<div class="info-item">';
            html += '<div class="info-label">Nama Siswa</div>';
            html += '<div class="info-value">' + studentName + '</div>';
            html += '</div>';
            html += '<div class="info-item">';
            html += '<div class="info-label">Asal Sekolah</div>';
            html += '<div class="info-value">' + studentSchool + '</div>';
            html += '</div>';
            html += '<div class="info-item">';
            html += '<div class="info-label">Skor Try Out</div>';
            html += '<div class="info-value">' + studentScore + '</div>';
            html += '</div>';
            html += '<div class="info-item">';
            html += '<div class="info-label">Tanggal</div>';
            html += '<div class="info-value">' + new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' }) + '</div>';
            html += '</div>';
            html += '</div>';
            html += '</div>';
            
            html += '<div class="choices-grid">';
            
            html += '<div class="choice-section">';
            html += '<div class="choice-title">PILIHAN 1</div>';
            if (choice1Programs.length > 0) {
                choice1Programs.forEach(function(program) {
                    html += generateProgramCard(program, 1);
                });
            } else {
                html += '<div class="empty-choice">Tidak ada pilihan</div>';
            }
            html += '</div>';
            
            html += '<div class="choice-section">';
            html += '<div class="choice-title">PILIHAN 2</div>';
            if (choice2Programs.length > 0) {
                choice2Programs.forEach(function(program) {
                    html += generateProgramCard(program, 2);
                });
            } else {
                html += '<div class="empty-choice">Tidak ada pilihan</div>';
            }
            html += '</div>';
            
            html += '</div>';
            
            html += '<div class="choices-grid">';
            
            html += '<div class="choice-section">';
            html += '<div class="choice-title">PILIHAN 3</div>';
            if (choice3Programs.length > 0) {
                choice3Programs.forEach(function(program) {
                    html += generateProgramCard(program, 3);
                });
            } else {
                html += '<div class="empty-choice">Tidak ada pilihan</div>';
            }
            html += '</div>';
            
            html += '<div class="choice-section">';
            html += '<div class="choice-title">PILIHAN 4</div>';
            if (choice4Programs.length > 0) {
                choice4Programs.forEach(function(program) {
                    html += generateProgramCard(program, 4);
                });
            } else {
                html += '<div class="empty-choice">Tidak ada pilihan</div>';
            }
            html += '</div>';
            
            html += '</div>';
            
            html += '<div class="signature-section">';
            html += '<div class="signature-box">';
            html += '<div class="signature-label">Siswa</div>';
            html += '<div class="signature-space"></div>';
            html += '<div class="signature-name">(' + studentName + ')</div>';
            html += '</div>';
            html += '<div class="signature-box">';
            html += '<div class="signature-label">Konsultan</div>';
            html += '<div class="signature-space"></div>';
            html += '<div class="signature-name">(' + getConsultantName() + ')</div>';
            html += '</div>';
            html += '</div>';
            
            html += '<div class="footer-section">';
            html += '<div class="footer-logo"><img src="https://iili.io/KsKbLep.png" alt="Logo Bimbel Attin"></div>';
            html += '<div class="footer-title">BIMBEL ATTIN</div>';
            html += '<div class="footer-text">Sistem Rekomendasi SNBT - Konsultasi Program Studi</div>';
            html += '<div class="footer-text">Untuk konsultasi lebih lanjut, hubungi Bimbel Attin</div>';
            html += '<div class="footer-date">Dicetak: ' + new Date().toLocaleString('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric', hour: '2-digit', minute: '2-digit' }) + '</div>';
            html += '</div>';
            
            html += '</div>';
            html += '</body></html>';
            
            return html;
        }
        
        function generateProgramCard(program, choiceNum) {
            const academicScore = Math.min((studentScore / program.pg) * 100, 100);
            const isNewProgram = program.applicants === 0 || program.applicants === null;
            const capacityCompetition = isNewProgram ? null : (program.capacity / program.applicants) * 100;
            const passChance = isNewProgram ? null : Math.min((academicScore * 0.6) + (capacityCompetition * 0.4), 100);
            
            let html = '<div class="program-card">';
            
            html += '<div style="display: flex; gap: 12px; margin-bottom: 10px; align-items: flex-start;">';
            html += '<div style="flex: 1;">';
            html += '<div class="program-header">';
            html += '<div class="program-header-left">';
            html += '<div class="university-name" style="margin-bottom: 0; font-size: 11px;">' + program.university + '</div>';
            html += '</div>';
            if (passChance !== null) {
                html += '<div class="peluang-box" style="min-width: 38px; padding: 2px 6px;">';
                html += '<div class="peluang-label" style="font-size: 6px;">Peluang</div>';
                html += '<div class="peluang-value" style="font-size: 12px;">' + Math.round(passChance) + '%</div>';
                html += '</div>';
            } else {
                html += '<div class="peluang-box new" style="min-width: 38px; padding: 2px 6px;">';
                html += '<div class="peluang-label" style="font-size: 6px;">Status</div>';
                html += '<div class="peluang-value" style="font-size: 9px;">BARU</div>';
                html += '</div>';
            }
            html += '</div>';
            html += '</div>';
            html += '</div>';
            
            html += '<div class="program-info">';
            if (program.majorCategory) {
                html += '<div class="program-value" style="font-size: 8px; opacity: 0.9;">' + program.majorCategory + '</div>';
            }
            html += '<div class="program-value">' + program.prodi + '</div>';
            html += '</div>';
            
            html += '<div class="badges-row">';
            html += '<span class="badge badge-level">' + program.level + '</span>';
            if (program.stream) {
                html += '<span class="badge badge-stream">' + program.stream + '</span>';
            }
            html += '<span class="badge badge-pg">PG: ' + formatPG(program.pg) + '</span>';
            html += '</div>';
            
            html += '<div class="stats-grid">';
            html += '<div class="stat-box">';
            html += '<div class="stat-label">PG</div>';
            html += '<div class="stat-value">' + formatPG(program.pg) + '</div>';
            html += '</div>';
            html += '<div class="stat-box">';
            html += '<div class="stat-label">Kapasitas</div>';
            html += '<div class="stat-value">' + program.capacity + '</div>';
            html += '</div>';
            html += '<div class="stat-box">';
            html += '<div class="stat-label">Peminat</div>';
            html += '<div class="stat-value">' + (isNewProgram ? 'Baru' : program.applicants) + '</div>';
            html += '</div>';
            html += '</div>';
            
            html += '</div>';
            
            return html;
        }

        function shareResults() {
            let allSelectedPrograms = [];
            for (const key in selectedChoices) {
                const choice = selectedChoices[key];
                if (choice !== null) {
                    if (Array.isArray(choice)) {
                        allSelectedPrograms = allSelectedPrograms.concat(choice);
                    } else {
                        allSelectedPrograms.push(choice);
                    }
                }
            }
            
            if (allSelectedPrograms.length === 0) {
                showToast('Tidak ada data untuk dibagikan', 'error');
                return;
            }
            
            const programToChoiceMap = new Map();
            if (selectedChoices.choice1 && Array.isArray(selectedChoices.choice1)) {
                selectedChoices.choice1.forEach(function(p) { programToChoiceMap.set(p, 1); });
            }
            if (selectedChoices.choice2 && Array.isArray(selectedChoices.choice2)) {
                selectedChoices.choice2.forEach(function(p) { programToChoiceMap.set(p, 2); });
            }
            if (selectedChoices.choice3 && Array.isArray(selectedChoices.choice3)) {
                selectedChoices.choice3.forEach(function(p) { programToChoiceMap.set(p, 3); });
            }
            if (selectedChoices.choice4 && Array.isArray(selectedChoices.choice4)) {
                selectedChoices.choice4.forEach(function(p) { programToChoiceMap.set(p, 4); });
            }
            
            let shareText = 'HASIL KONSULTASI SNBT - BIMBEL ATTIN\n\n';
            shareText += 'Nama: ' + studentName + '\n';
            shareText += 'Sekolah: ' + studentSchool + '\n';
            shareText += 'Skor Try Out: ' + studentScore + '\n';
            shareText += 'Tanggal: ' + new Date().toLocaleDateString('id-ID') + '\n\n';
            shareText += '===================================\n\n';
            shareText += 'REKOMENDASI PROGRAM STUDI:\n\n';
            
            allSelectedPrograms.forEach(function(program, index) {
                const choiceNum = programToChoiceMap.get(program) || 1;
                const scorePercentage = Math.min((studentScore / program.pg) * 100, 100);
                const isNewProgram = program.applicants === 0 || program.applicants === null;
                const competitionPercentage = isNewProgram ? null : (program.capacity / program.applicants) * 100;
                const passChance = isNewProgram ? null : Math.min((scorePercentage * 0.6) + (competitionPercentage * 0.4), 100);
                
                shareText += (index + 1) + '. PILIHAN ' + choiceNum + '\n';
                shareText += 'Universitas: ' + program.university + '\n';
                if (program.majorCategory) {
                    shareText += 'Kategori: ' + program.majorCategory + '\n';
                }
                shareText += 'Prodi: ' + program.prodi + '\n';
                shareText += 'Jenjang: ' + program.level;
                if (program.stream) {
                    shareText += ' | ' + program.stream;
                }
                shareText += '\n';
                if (passChance !== null) {
                    shareText += 'PG: ' + formatPG(program.pg) + ' | Peluang: ' + Math.round(passChance) + '%\n';
                } else {
                    shareText += 'PG: ' + formatPG(program.pg) + ' | Peluang: Program Baru\n';
                }
                shareText += 'Daya Tampung: ' + program.capacity + ' | Peminat: ' + (isNewProgram ? 'Program Baru' : program.applicants) + '\n';
                shareText += '\n';
            });
            
            shareText += '===================================\n';
            shareText += 'Bimbel Attin - Sistem Rekomendasi SNBT\n';
            shareText += 'Konsultasi lebih lanjut: hubungi Bimbel Attin';
            
            if (navigator.share) {
                navigator.share({
                    title: 'Hasil Konsultasi SNBT - ' + studentName,
                    text: shareText
                }).then(function() {
                    showToast('Hasil berhasil dibagikan!', 'success');
                }).catch(function(error) {
                    if (error.name !== 'AbortError') {
                        copyToClipboardFallback(shareText);
                    }
                });
            } else {
                copyToClipboardFallback(shareText);
            }
        }
        
        function copyToClipboardFallback(text) {
            const textarea = document.createElement('textarea');
            textarea.value = text;
            textarea.style.position = 'fixed';
            textarea.style.top = '0';
            textarea.style.left = '0';
            textarea.style.opacity = '0';
            document.body.appendChild(textarea);
            textarea.focus();
            textarea.select();
            
            try {
                const successful = document.execCommand('copy');
                if (successful) {
                    showToast('Hasil disalin ke clipboard!', 'success');
                } else {
                    showShareModal(text);
                }
            } catch (err) {
                showShareModal(text);
            }
            
            document.body.removeChild(textarea);
        }
        
        function showShareModal(text) {
            const modal = document.createElement('div');
            modal.className = 'modal-overlay';
            modal.style.display = 'flex';
            modal.innerHTML = '<div class="modal-content"><div class="glass-card">' +
                '<h2 class="text-white text-xl font-bold mb-4" style="text-align: center;">Salin Hasil Konsultasi</h2>' +
                '<textarea readonly style="width: 100%; height: 300px; padding: 15px; border-radius: 12px; background: rgba(0,0,0,0.5); color: white; border: 2px solid rgba(220, 38, 38, 0.3); font-size: 12px; font-family: monospace; margin-bottom: 15px; resize: vertical;">' + text + '</textarea>' +
                '<button class="glass-button" onclick="this.closest(\'.modal-overlay\').remove();">Tutup</button>' +
                '</div></div>';
            document.body.appendChild(modal);
            
            const textarea = modal.querySelector('textarea');
            textarea.select();
            
            showToast('Silakan salin teks di atas', 'info');
        }

        function viewPassingGrades(choiceNum) {
            if (event) event.stopPropagation();
            
            if (masterData.length === 0) {
                showToast('Silakan upload master data terlebih dahulu', 'error');
                return;
            }

            let allData = [];
            let thresholdScore = studentScore;
            const hasScore = studentScore && studentScore > 0;
            
            // Define level filter based on choice
            let levelFilter = [];
            if (choiceNum === 1) {
                levelFilter = ['S1'];
            } else if (choiceNum === 2) {
                if (!hasScore) {
                    showToast('Masukkan skor Try Out terlebih dahulu', 'error');
                    return;
                }
                thresholdScore = studentScore;
                levelFilter = ['S1', 'D4'];
            } else if (choiceNum === 3) {
                if (!hasScore) {
                    showToast('Masukkan skor Try Out terlebih dahulu', 'error');
                    return;
                }
                thresholdScore = studentScore - 50;
                levelFilter = ['D4', 'D3'];
            } else if (choiceNum === 4) {
                if (!hasScore) {
                    showToast('Masukkan skor Try Out terlebih dahulu', 'error');
                    return;
                }
                thresholdScore = studentScore - 100;
                levelFilter = ['D3'];
            }

            // Filter by level only - show ALL programs regardless of PG vs threshold
            allData = masterData.filter(function(d) {
                return levelFilter.indexOf(d.level) !== -1;
            });

            // Sort by PG descending (highest to lowest)
            allData.sort(function(a, b) { return b.pg - a.pg; });

            const pgModal = document.getElementById('pg-modal');
            const pgList = document.getElementById('pg-list');
            const pgTitle = document.getElementById('pg-modal-title');

            pgTitle.textContent = 'Passing Grade - Pilihan ' + choiceNum;
            pgList.innerHTML = '';

            if (allData.length === 0) {
                pgList.innerHTML = '<div class="glass-card" style="text-align: center; padding: 30px;"><p class="text-white" style="opacity: 0.6;">Tidak ada program studi untuk pilihan ini</p></div>';
                pgModal.style.display = 'flex';
                return;
            }

            // Count recommended and not recommended
            let recommendedCount = 0;
            let notRecommendedCount = 0;
            
            allData.forEach(function(program) {
                if (program.pg <= thresholdScore) {
                    recommendedCount++;
                } else {
                    notRecommendedCount++;
                }
            });

            const summaryCard = document.createElement('div');
            summaryCard.className = 'glass-card';
            summaryCard.style.marginBottom = '15px';
            summaryCard.style.background = 'linear-gradient(135deg, rgba(102, 126, 234, 0.3), rgba(118, 75, 162, 0.3))';
            
            let summaryHtml = '<div style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px;">';
            summaryHtml += '<div style="text-align: center; padding: 10px; background: rgba(6, 214, 160, 0.3); border-radius: 10px;">';
            summaryHtml += '<p class="text-white" style="font-size: 12px; opacity: 0.8; margin-bottom: 4px;">Direkomendasikan</p>';
            summaryHtml += '<p class="text-white" style="font-size: 18px; font-weight: bold; color: #06d6a0;">' + recommendedCount + '</p>';
            summaryHtml += '</div>';
            summaryHtml += '<div style="text-align: center; padding: 10px; background: rgba(245, 87, 108, 0.3); border-radius: 10px;">';
            summaryHtml += '<p class="text-white" style="font-size: 12px; opacity: 0.8; margin-bottom: 4px;">Tidak Direkomendasikan</p>';
            summaryHtml += '<p class="text-white" style="font-size: 18px; font-weight: bold; color: #f5576c;">' + notRecommendedCount + '</p>';
            summaryHtml += '</div>';
            summaryHtml += '</div>';
            
            summaryCard.innerHTML = summaryHtml;
            pgList.appendChild(summaryCard);

            // Display all programs
            allData.forEach(function(program, index) {
                const isRecommended = program.pg <= thresholdScore;
                const card = document.createElement('div');
                card.className = 'glass-card';
                card.style.marginBottom = '12px';
                card.style.padding = '15px';
                card.style.opacity = isRecommended ? '1' : '0.7';
                card.style.borderColor = isRecommended ? 'rgba(220, 38, 38, 0.4)' : 'rgba(245, 87, 108, 0.5)';

                let html = '<div style="margin-bottom: 8px;">';
                html += '<div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 8px;">';
                html += '<div style="flex: 1;">';
                html += '<h4 class="text-white font-bold" style="font-size: 15px; margin-bottom: 4px; text-transform: uppercase; letter-spacing: 0.3px;">' + program.university + '</h4>';
                if (program.majorCategory) {
                    html += '<p class="text-white" style="opacity: 0.85; font-size: 13px; margin-bottom: 3px;">' + program.majorCategory + '</p>';
                }
                html += '<p class="text-white" style="opacity: 0.9; font-size: 14px; margin-bottom: 6px; font-weight: 600;">' + program.prodi + '</p>';
                html += '</div>';
                
                if (isRecommended) {
                    html += '<div style="background: linear-gradient(135deg, #06d6a0, #11998e); padding: 6px 14px; border-radius: 8px; text-align: center;">';
                    html += '<p class="text-white" style="font-size: 11px; opacity: 0.9; margin-bottom: 2px;">PG</p>';
                    html += '<p class="text-white" style="font-size: 18px; font-weight: bold;">' + formatPG(program.pg) + '</p>';
                    html += '</div>';
                } else {
                    html += '<div style="background: linear-gradient(135deg, #f5576c, #f093fb); padding: 6px 14px; border-radius: 8px; text-align: center;">';
                    html += '<p class="text-white" style="font-size: 11px; opacity: 0.9; margin-bottom: 2px;">TIDAK</p>';
                    html += '<p class="text-white" style="font-size: 12px; font-weight: bold;">DIREK.</p>';
                    html += '</div>';
                }
                html += '</div>';

                html += '<div style="display: flex; gap: 6px; flex-wrap: wrap; margin-bottom: 8px;">';
                html += '<span class="status-badge" style="background: rgba(102, 126, 234, 0.4); color: white;">' + program.level + '</span>';
                if (program.stream) {
                    html += '<span class="status-badge" style="background: rgba(220, 38, 38, 0.4); color: white;">' + program.stream + '</span>';
                }
                html += '<span class="status-badge" style="background: rgba(6, 214, 160, 0.4); color: white;">PG: ' + formatPG(program.pg) + '</span>';
                html += '</div>';

                if (!isRecommended) {
                    const diff = Math.round((program.pg - thresholdScore) * 100) / 100;
                    html += '<div style="background: rgba(245, 87, 108, 0.2); border-left: 3px solid #f5576c; padding: 8px 10px; border-radius: 6px;">';
                    html += '<p class="text-white" style="font-size: 12px; color: #fecaca;"><strong>⚠️ Melebihi ' + diff + ' poin</strong> dari threshold pilihan ' + choiceNum + ' (' + thresholdScore + ')</p>';
                    html += '</div>';
                }

                html += '</div>';

                card.innerHTML = html;
                pgList.appendChild(card);
            });

            pgModal.style.display = 'flex';
        }

        function closePGModal() {
            const pgModal = document.getElementById('pg-modal');
            if (pgModal) pgModal.style.display = 'none';
        }

        function startNewConsultation() {
            studentName = '';
            studentSchool = '';
            studentScore = 0;
            
            document.getElementById('student-name').value = '';
            document.getElementById('student-school').value = '';
            document.getElementById('student-score').value = '';
            
            selectedChoices = {
                choice1: null,
                choice2: null,
                choice3: null,
                choice4: null
            };
            
            for (let i = 1; i <= 4; i++) {
                const checkmark = document.getElementById('check-' + i);
                const card = document.getElementById('choice-card-' + i);
                if (checkmark) checkmark.style.display = 'none';
                if (card) card.classList.remove('selected');
            }
            
            const analyzeBtn = document.getElementById('analyze-button');
            if (analyzeBtn) analyzeBtn.style.display = 'none';
            
            const choicesStatus = document.getElementById('choices-status');
            if (choicesStatus) choicesStatus.textContent = 'Pilih program studi untuk setiap pilihan';
            
            showToast('Memulai konsultasi baru', 'success');
            showPage('page-3');
        }

        function loadUsersList() {
            const usersList = document.getElementById('users-list');
            usersList.innerHTML = '';
            
            if (Object.keys(registeredUsers).length === 0) {
                usersList.innerHTML = '<div class="glass-card" style="text-align: center; padding: 20px;"><p class="text-white" style="opacity: 0.6;">Belum ada user terdaftar</p></div>';
                return;
            }
            
            const fragment = document.createDocumentFragment();
            
            // Add admin user display
            const adminItem = document.createElement('div');
            adminItem.className = 'user-item';
            adminItem.innerHTML = '<div class="user-info"><div class="user-username">attin <span class="admin-badge">Admin</span></div></div><div style="width: 80px;"></div>';
            fragment.appendChild(adminItem);
            
            // Add registered users
            for (const username in registeredUsers) {
                const userItem = document.createElement('div');
                userItem.className = 'user-item';
                const password = registeredUsers[username];
                const passwordMasked = '•'.repeat(password.length);
                const uniqueId = 'pwd_' + Math.random().toString(36).substr(2, 9);
                
                userItem.innerHTML = '<div class="user-info"><div class="user-username">' + username + '</div><div class="user-password"><span class="password-text" id="' + uniqueId + '">' + passwordMasked + '</span><button type="button" class="toggle-password-btn" onclick="togglePasswordVisibility(\'' + uniqueId + '\', \'' + password.replace(/'/g, "\\'") + '\', this)">Lihat</button></div></div><div style="display: flex; gap: 8px;"><button class="delete-user-btn" onclick="editUser(\'' + username.replace(/'/g, "\\'") + '\')" style="background: linear-gradient(135deg, #667eea, #764ba2);">Edit</button><button class="delete-user-btn" onclick="deleteUser(\'' + username.replace(/'/g, "\\'") + '\')">Hapus</button></div>';
                fragment.appendChild(userItem);
            }
            
            usersList.appendChild(fragment);
        }

        async function addNewUser() {
            const newUsername = document.getElementById('new-username').value.trim();
            const newPassword = document.getElementById('new-password').value.trim();
            
            if (!newUsername || !newPassword) {
                showToast('Silakan isi username dan password', 'error');
                return;
            }
            
            if (newUsername.toLowerCase() === ADMIN_USERNAME) {
                showToast('Username tidak boleh sama dengan admin', 'error');
                return;
            }
            
            if (registeredUsers[newUsername]) {
                showToast('Username sudah terdaftar', 'error');
                return;
            }
            
            if (newUsername.length < 3) {
                showToast('Username minimal 3 karakter', 'error');
                return;
            }
            
            if (newPassword.length < 4) {
                showToast('Password minimal 4 karakter', 'error');
                return;
            }
            
            showLoading();
            
            // Simpan ke backend jika tidak menggunakan local API
            if (!USE_LOCAL_API) {
                const backendSuccess = await saveUserToBackend(newUsername, newPassword, 'add');
                if (!backendSuccess) {
                    hideLoading();
                    showToast('Gagal menyimpan ke backend. Coba lagi.', 'error');
                    return;
                }
            }
            
            registeredUsers[newUsername] = newPassword;
            saveRegisteredUsers();
            
            document.getElementById('new-username').value = '';
            document.getElementById('new-password').value = '';
            
            hideLoading();
            loadUsersList();
            showToast('User ' + newUsername + ' berhasil ditambahkan! ✓', 'success');
        }

        function deleteUser(username) {
            if (!registeredUsers[username]) {
                showToast('User tidak ditemukan', 'error');
                return;
            }
            
            showDeleteConfirmation(username);
        }

        function showDeleteConfirmation(username) {
            const modal = document.createElement('div');
            modal.className = 'modal-overlay';
            modal.style.display = 'flex';
            modal.innerHTML = '<div class="modal-content"><div class="glass-card">' +
                '<div style="text-align: center; margin-bottom: 20px;">' +
                '<div style="width: 70px; height: 70px; background: linear-gradient(135deg, #f5576c, #f093fb); border-radius: 50%; margin: 0 auto 15px; display: flex; align-items: center; justify-content: center; box-shadow: 0 8px 20px rgba(245, 87, 108, 0.4);">' +
                '<svg width="35" height="35" viewbox="0 0 24 24" fill="none" stroke="white" stroke-width="2"><polyline points="3 6 5 4 21 4 23 6 23 20a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2V6"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>' +
                '</div>' +
                '<h2 class="text-white text-xl font-bold mb-2">Hapus User?</h2>' +
                '<p class="text-white" style="opacity: 0.8; font-size: 14px;">Anda yakin ingin menghapus user <strong>' + username + '</strong>? Tindakan ini tidak dapat dibatalkan.</p>' +
                '</div>' +
                '<div style="display: grid; gap: 12px;">' +
                '<button class="glass-button" onclick="confirmDeleteUser(\'' + username.replace(/'/g, "\\'") + '\')" style="background: linear-gradient(135deg, #f5576c, #f093fb);">' +
                '<svg width="20" height="20" viewbox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="display: inline-block; margin-right: 8px; vertical-align: middle;"><polyline points="3 6 5 4 21 4 23 6 23 20a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2V6"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg> Ya, Hapus User </button>' +
                '<button class="glass-button" onclick="this.closest(\'.modal-overlay\').remove();" style="background: linear-gradient(135deg, #4b5563, #374151);"> Batal </button>' +
                '</div>' +
                '</div></div>';
            document.body.appendChild(modal);
        }

        async function confirmDeleteUser(username) {
            showLoading();
            
            // Hapus dari backend jika tidak menggunakan local API
            if (!USE_LOCAL_API) {
                const backendSuccess = await saveUserToBackend(username, '', 'delete');
                if (!backendSuccess) {
                    hideLoading();
                    showToast('Gagal menghapus dari backend. Coba lagi.', 'error');
                    return;
                }
            }
            
            delete registeredUsers[username];
            saveRegisteredUsers();
            
            document.querySelectorAll('.modal-overlay').forEach(function(m) { m.remove(); });
            
            hideLoading();
            loadUsersList();
            showToast('User ' + username + ' berhasil dihapus! ✓', 'success');
        }

        function editUser(username) {
            const password = registeredUsers[username];
            
            const modal = document.createElement('div');
            modal.className = 'modal-overlay';
            modal.style.display = 'flex';
            modal.innerHTML = '<div class="modal-content"><div class="glass-card">' +
                '<h2 class="text-white text-xl font-bold mb-4" style="text-align: center;">Edit User</h2>' +
                '<div style="margin-bottom: 15px;">' +
                '<label for="edit-username-input" class="text-white block mb-2" style="font-size: 14px;">Username</label>' +
                '<input type="text" class="glass-input" id="edit-username-input" value="' + username + '" disabled style="opacity: 0.6;">' +
                '</div>' +
                '<div style="margin-bottom: 20px;">' +
                '<label for="edit-password-input" class="text-white block mb-2" style="font-size: 14px;">Password Baru</label>' +
                '<input type="password" class="glass-input" id="edit-password-input" placeholder="Masukkan password baru (kosongkan jika tidak diubah)">' +
                '</div>' +
                '<div style="display: grid; gap: 12px;">' +
                '<button class="glass-button" onclick="saveEditUser(\'' + username.replace(/'/g, "\\'") + '\')" style="background: linear-gradient(135deg, #06d6a0, #11998e);">' +
                '<svg width="20" height="20" viewbox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="display: inline-block; margin-right: 8px; vertical-align: middle;"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg> Simpan Perubahan </button>' +
                '<button class="glass-button" onclick="this.closest(\'.modal-overlay\').remove();" style="background: linear-gradient(135deg, #4b5563, #374151);"> Batal </button>' +
                '</div>' +
                '</div></div>';
            document.body.appendChild(modal);
            
            setTimeout(function() {
                const input = document.getElementById('edit-password-input');
                if (input) input.focus();
            }, 100);
        }

        async function saveEditUser(username) {
            const newPassword = document.getElementById('edit-password-input').value.trim();
            
            if (!newPassword) {
                showToast('Masukkan password baru', 'error');
                return;
            }
            
            if (newPassword.length < 4) {
                showToast('Password minimal 4 karakter', 'error');
                return;
            }
            
            showLoading();
            
            // Update di backend jika tidak menggunakan local API
            if (!USE_LOCAL_API) {
                const backendSuccess = await saveUserToBackend(username, newPassword, 'update');
                if (!backendSuccess) {
                    hideLoading();
                    showToast('Gagal memperbarui di backend. Coba lagi.', 'error');
                    return;
                }
            }
            
            registeredUsers[username] = newPassword;
            saveRegisteredUsers();
            
            document.querySelectorAll('.modal-overlay').forEach(function(m) { m.remove(); });
            
            hideLoading();
            loadUsersList();
            showToast('Password user ' + username + ' berhasil diperbarui! ✓', 'success');
        }

        function togglePasswordVisibility(elementId, password, button) {
            const element = document.getElementById(elementId);
            if (!element) return;
            
            if (element.textContent === password) {
                element.textContent = '•'.repeat(password.length);
                button.textContent = 'Lihat';
            } else {
                element.textContent = password;
                button.textContent = 'Sembunyikan';
            }
        }

        function editAdminCredentials() {
            const modal = document.createElement('div');
            modal.className = 'modal-overlay';
            modal.style.display = 'flex';
            
            let adminUsernameValue = ADMIN_USERNAME;
            let adminPasswordValue = ADMIN_PASSWORD;
            let showPassword = false;
            
            modal.innerHTML = '<div class="modal-content"><div class="glass-card">' +
                '<div style="text-align: center; margin-bottom: 20px;">' +
                '<div style="width: 70px; height: 70px; background: linear-gradient(135deg, #f093fb, #f5576c); border-radius: 50%; margin: 0 auto 15px; display: flex; align-items: center; justify-content: center; box-shadow: 0 8px 20px rgba(240, 147, 251, 0.4);">' +
                '<svg width="35" height="35" viewbox="0 0 24 24" fill="none" stroke="white" stroke-width="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>' +
                '</div>' +
                '<h2 class="text-white text-xl font-bold mb-2">Edit Admin Credentials</h2>' +
                '<p class="text-white" style="opacity: 0.8; font-size: 13px;">Ubah username dan password admin</p>' +
                '</div>' +
                '<div style="margin-bottom: 15px;">' +
                '<label for="admin-username-input" class="text-white block mb-2" style="font-size: 14px;">Username Admin</label>' +
                '<input type="text" class="glass-input" id="admin-username-input" value="' + adminUsernameValue + '" placeholder="Username admin">' +
                '</div>' +
                '<div style="margin-bottom: 20px;">' +
                '<label for="admin-password-input" class="text-white block mb-2" style="font-size: 14px;">Password Admin</label>' +
                '<div style="display: flex; gap: 8px; align-items: center;">' +
                '<input type="password" class="glass-input" id="admin-password-input" value="' + adminPasswordValue + '" placeholder="Password admin" style="flex: 1;">' +
                '<button type="button" class="toggle-password-btn" id="toggle-admin-pwd" onclick="toggleAdminPasswordField()" style="padding: 8px 12px; margin-top: 0;">Lihat</button>' +
                '</div>' +
                '</div>' +
                '<div style="display: grid; gap: 12px;">' +
                '<button class="glass-button" onclick="saveAdminCredentials()" style="background: linear-gradient(135deg, #06d6a0, #11998e);">' +
                '<svg width="20" height="20" viewbox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="display: inline-block; margin-right: 8px; vertical-align: middle;"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg> Simpan Perubahan </button>' +
                '<button class="glass-button" onclick="this.closest(\'.modal-overlay\').remove();" style="background: linear-gradient(135deg, #4b5563, #374151);"> Batal </button>' +
                '</div>' +
                '</div></div>';
            document.body.appendChild(modal);
            
            setTimeout(function() {
                const input = document.getElementById('admin-username-input');
                if (input) input.focus();
            }, 100);
        }

        function toggleAdminPasswordField() {
            const input = document.getElementById('admin-password-input');
            const btn = document.getElementById('toggle-admin-pwd');
            
            if (input.type === 'password') {
                input.type = 'text';
                btn.textContent = 'Sembunyikan';
            } else {
                input.type = 'password';
                btn.textContent = 'Lihat';
            }
        }

        async function saveAdminCredentials() {
            const newUsername = document.getElementById('admin-username-input').value.trim();
            const newPassword = document.getElementById('admin-password-input').value.trim();
            
            if (!newUsername || !newPassword) {
                showToast('Username dan password tidak boleh kosong', 'error');
                return;
            }
            
            if (newUsername.length < 3) {
                showToast('Username minimal 3 karakter', 'error');
                return;
            }
            
            if (newPassword.length < 4) {
                showToast('Password minimal 4 karakter', 'error');
                return;
            }
            
            showLoading();
            
            // Update admin di backend jika tidak menggunakan local API
            if (!USE_LOCAL_API) {
                try {
                    const response = await fetch(API_BASE_URL + '/admin/credentials', {
                        method: 'PUT',
                        headers: {
                            'Authorization': 'Bearer ' + API_KEY,
                            'Content-Type': 'application/json'
                        },
                        body: JSON.stringify({
                            username: newUsername,
                            password: newPassword
                        })
                    });

                    if (!response.ok) {
                        throw new Error('Backend error: ' + response.status);
                    }

                    const result = await response.json();
                    if (!result.success) {
                        throw new Error(result.message || 'Failed to update admin credentials');
                    }
                } catch (error) {
                    hideLoading();
                    showToast('Gagal memperbarui di backend: ' + error.message, 'error');
                    return;
                }
            }
            
            // Update admin credentials globally
            ADMIN_USERNAME = newUsername;
            ADMIN_PASSWORD = newPassword;
            
            // Save to storage
            Storage.set('admin_credentials', {
                username: newUsername,
                password: newPassword
            });
            
            document.querySelectorAll('.modal-overlay').forEach(function(m) { m.remove(); });
            
            hideLoading();
            showToast('Admin credentials berhasil diperbarui! ✓', 'success');
        }
    </script>
 <script>(function(){function c(){var b=a.contentDocument||a.contentWindow.document;if(b){var d=b.createElement('script');d.innerHTML="window.__CF$cv$params={r:'9fc3362a80a4c691',t:'MTc3ODg1OTMxNy4wMDAwMDA='};var a=document.createElement('script');a.nonce='';a.src='/cdn-cgi/challenge-platform/scripts/jsd/main.js';document.getElementsByTagName('head')[0].appendChild(a);";b.getElementsByTagName('head')[0].appendChild(d)}}if(document.body){var a=document.createElement('iframe');a.height=1;a.width=1;a.style.position='absolute';a.style.top=0;a.style.left=0;a.style.border='none';a.style.visibility='hidden';document.body.appendChild(a);if('loading'!==document.readyState)c();else if(window.addEventListener)document.addEventListener('DOMContentLoaded',c);else{var e=document.onreadystatechange||function(){};document.onreadystatechange=function(b){e(b);'loading'!==document.readyState&&(document.onreadystatechange=e,c())}}}})();</script></body>
</html>
