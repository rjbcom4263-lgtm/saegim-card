importScripts('https://www.gstatic.com/firebasejs/11.6.0/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/11.6.0/firebase-messaging-compat.js');

firebase.initializeApp({
  apiKey:"AIzaSyA8uBfMGkg9gLiHiqnFh_Klmd0etvk0qQg",
  authDomain:"saegim-memory.firebaseapp.com",
  projectId:"saegim-memory",
  storageBucket:"saegim-memory.firebasestorage.app",
  messagingSenderId:"888229102391",
  appId:"1:888229102391:web:f4fe41ece07ddacd2a10b2"
});

const messaging=firebase.messaging();

messaging.onBackgroundMessage(function(payload){
  const n=payload.notification||{};
  const data=payload.data||{};
  self.registration.showNotification(n.title||'새김 알림',{
    body:n.body||'새 메시지가 도착했습니다.',
    icon:'/icon-192.png',
    badge:'/icon-96.png',
    data:{url:data.url||'/'}
  });
});

self.addEventListener('notificationclick',function(event){
  event.notification.close();
  const url=(event.notification.data&&event.notification.data.url)||'/';
  event.waitUntil(clients.openWindow(url));
});
