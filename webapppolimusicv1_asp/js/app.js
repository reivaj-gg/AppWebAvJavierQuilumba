/* ================================================================
   PoliMusic - Main JavaScript
   Audio Player, CRUD AJAX, UI Interactions
   ================================================================ */

(function() {
    'use strict';

    // --- Global State ---
    var playlist = [];       // Array of {id, title, path, plays}
    var currentIndex = -1;    // Currently loaded song index in playlist
    var isPlaying = false;

    // --- Audio Element ---
    var audio = new Audio();
    audio.preload = 'metadata';

    // --- DOM Elements (cached) ---
    var playerBar = document.querySelector('.player-bar');
    var playerTitle = document.querySelector('.player-title');
    var playerArtist = document.querySelector('.player-artist');
    var playerCover = document.querySelector('.player-cover');
    var playBtn = document.getElementById('playBtn');
    var prevBtn = document.getElementById('prevBtn');
    var nextBtn = document.getElementById('nextBtn');
    var progressFill = document.getElementById('progressFill');
    var progressBar = document.getElementById('progressBar');
    var currentTimeEl = document.getElementById('currentTime');
    var durationEl = document.getElementById('duration');
    var volumeFill = document.getElementById('volumeFill');
    var volumeBar = document.getElementById('volumeBar');
    var muteBtn = document.getElementById('muteBtn');

    // ==================================================================
    // INITIALIZATION
    // ==================================================================

    function init() {
        loadPlaylist();
        bindAudioEvents();
        bindPlayerControls();
        bindKeyboardShortcuts();
        setVolume(0.7);
    }

    // Load songs from the data attributes rendered by index.asp
    // (supports both the table rows and the legacy card layout)
    function loadPlaylist() {
        var items = document.querySelectorAll('.song-row, .song-card');

        items.forEach(function(item) {
            var songId = item.getAttribute('data-song-id');
            var songTitle = item.getAttribute('data-song-title');
            var songPath = item.getAttribute('data-song-path');
            var songPlays = item.getAttribute('data-song-plays') || '0';
            var songArtist = item.getAttribute('data-song-artist') || 'Unknown Artist';

            if (!songId || !songPath) return;

            playlist.push({
                id: songId,
                title: songTitle,
                path: songPath,
                artist: songArtist,
                plays: parseInt(songPlays, 10) || 0
            });

            // Click anywhere on the row / card to play
            item.addEventListener('click', function(e) {
                if (e.target.closest('.row-actions') || e.target.closest('.song-card-actions')) return;
                e.preventDefault();
                playSongByIndex(getPlaylistIndex(songId));
            });

            // Keyboard support (Enter / Space on the focused row)
            item.addEventListener('keydown', function(e) {
                if (e.key === 'Enter' || e.key === ' ') {
                    e.preventDefault();
                    playSongByIndex(getPlaylistIndex(songId));
                }
            });

            // Dedicated play button inside the row / card
            var innerPlay = item.querySelector('.row-play, .song-card-play-btn');
            if (innerPlay) {
                innerPlay.addEventListener('click', function(e) {
                    e.stopPropagation();
                    e.preventDefault();
                    playSongByIndex(getPlaylistIndex(songId));
                });
            }
        });
    }

    // Exposed so inline handlers can request playback by song id
    window.playRow = function(songId) {
        playSongByIndex(getPlaylistIndex(songId));
    };

    function getPlaylistIndex(songId) {
        for (var i = 0; i < playlist.length; i++) {
            if (playlist[i].id == songId) return i;
        }
        return -1;
    }

    // ==================================================================
    // AUDIO EVENT LISTENERS
    // ==================================================================

    function bindAudioEvents() {
        audio.addEventListener('loadedmetadata', function() {
            durationEl.textContent = formatTime(audio.duration);
        });

        audio.addEventListener('timeupdate', function() {
            if (audio.duration > 0) {
                var pct = (audio.currentTime / audio.duration) * 100;
                progressFill.style.width = pct + '%';
                currentTimeEl.textContent = formatTime(audio.currentTime);
            }
        });

        audio.addEventListener('ended', function() {
            nextSong();
        });

        audio.addEventListener('play', function() {
            isPlaying = true;
            updatePlayButton();
        });

        audio.addEventListener('pause', function() {
            isPlaying = false;
            updatePlayButton();
        });

        audio.addEventListener('error', function() {
            showToast('Error loading audio file. Check the song path.', 'error');
            isPlaying = false;
            updatePlayButton();
        });

        audio.addEventListener('volumechange', function() {
            updateVolumeUI();
        });
    }

    // ==================================================================
    // PLAYER CONTROLS
    // ==================================================================

    function bindPlayerControls() {
        if (playBtn) {
            playBtn.addEventListener('click', togglePlay);
        }
        if (prevBtn) {
            prevBtn.addEventListener('click', prevSong);
        }
        if (nextBtn) {
            nextBtn.addEventListener('click', nextSong);
        }
        if (progressBar) {
            progressBar.addEventListener('click', seek);
        }
        if (volumeBar) {
            volumeBar.addEventListener('click', setVolumeFromBar);
        }
        if (muteBtn) {
            muteBtn.addEventListener('click', toggleMute);
        }
    }

    function togglePlay() {
        if (currentIndex === -1) {
            if (playlist.length > 0) {
                playSongByIndex(0);
            } else {
                showToast('No songs available to play', 'info');
            }
            return;
        }
        if (isPlaying) {
            audio.pause();
        } else {
            audio.play().catch(function() {
                showToast('Unable to play audio', 'error');
            });
        }
    }

    function playSongByIndex(index) {
        if (index < 0 || index >= playlist.length) return;

        currentIndex = index;
        var song = playlist[index];

        // Build path - resolve relative paths from DB
        var songUrl = song.path;
        // DB stores paths like ~/songFiles/bensound-adventure.mp3
        // We resolve to SongFilesAndImages_resources/SongFiles/ in web root
        songUrl = songUrl.replace('~/', '');
        songUrl = songUrl.replace('songFiles/', 'SongFilesAndImages_resources/SongFiles/');
        songUrl = songUrl.replace('songfiles/', 'SongFilesAndImages_resources/SongFiles/');

        audio.src = songUrl;
        audio.play().then(function() {
            // Update UI
            updatePlayerInfo(song);
            highlightActiveCard(song.id);
            // Increment plays via AJAX
            incrementPlays(song.id);
        }).catch(function(err) {
            showToast('Could not play: ' + song.title, 'error');
            console.error('Audio play error:', err);
        });
    }

    function prevSong() {
        if (currentIndex > 0) {
            playSongByIndex(currentIndex - 1);
        } else if (playlist.length > 0) {
            playSongByIndex(playlist.length - 1);
        }
    }

    function nextSong() {
        if (currentIndex < playlist.length - 1) {
            playSongByIndex(currentIndex + 1);
        } else if (playlist.length > 0) {
            // Stop at end (or loop)
            playSongByIndex(0);
        }
    }

    function updatePlayButton() {
        if (!playBtn) return;
        var icon = playBtn.querySelector('svg');
        if (isPlaying) {
            icon.innerHTML = '<path d="M6 4h4v16H6zM14 4h4v16h-4z"/>';
        } else {
            icon.innerHTML = '<path d="M8 5v14l11-7z"/>';
        }
    }

    function updatePlayerInfo(song) {
        if (playerTitle) playerTitle.textContent = song.title;
        if (playerArtist) playerArtist.textContent = song.artist || 'Unknown Artist';
        if (playerCover) {
            playerCover.innerHTML = '<span class="cover-placeholder">' + song.title.charAt(0).toUpperCase() + '</span>';
        }
        if (playerBar) {
            playerBar.classList.remove('empty');
        }
    }

    function highlightActiveCard(songId) {
        document.querySelectorAll('.song-row, .song-card').forEach(function(item) {
            item.classList.remove('playing');
            if (item.getAttribute('data-song-id') == songId) {
                item.classList.add('playing');
            }
        });
    }

    function seek(e) {
        if (audio.duration > 0) {
            var rect = progressBar.getBoundingClientRect();
            var pct = (e.clientX - rect.left) / rect.width;
            audio.currentTime = audio.duration * pct;
        }
    }

    // ==================================================================
    // VOLUME CONTROLS
    // ==================================================================

    function setVolume(vol) {
        audio.volume = Math.max(0, Math.min(1, vol));
        updateVolumeUI();
    }

    function setVolumeFromBar(e) {
        var rect = volumeBar.getBoundingClientRect();
        var pct = (e.clientX - rect.left) / rect.width;
        setVolume(pct);
    }

    function updateVolumeUI() {
        if (volumeFill) {
            volumeFill.style.width = (audio.volume * 100) + '%';
        }
    }

    var prevVolume = 0.7;
    function toggleMute() {
        if (audio.volume > 0) {
            prevVolume = audio.volume;
            setVolume(0);
        } else {
            setVolume(prevVolume);
        }
    }

    // ==================================================================
    // AJAX: Increment play count
    // ==================================================================

    function incrementPlays(songId) {
        var xhr = new XMLHttpRequest();
        xhr.open('POST', 'api/increment_plays.asp', true);
        xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        xhr.onreadystatechange = function() {
            if (xhr.readyState !== 4 || xhr.status !== 200) return;

            var newCount = null;
            try {
                var data = JSON.parse(xhr.responseText);
                if (data && data.success && typeof data.plays !== 'undefined') {
                    newCount = data.plays;
                }
            } catch (err) { /* keep the local increment below */ }

            var row = document.querySelector('[data-song-id="' + songId + '"]');
            if (!row) return;

            var playsEl = row.querySelector('.plays-count');
            if (!playsEl) return;

            if (newCount === null) {
                newCount = (parseInt(playsEl.textContent.replace(/\D/g, ''), 10) || 0) + 1;
            }
            playsEl.textContent = newCount;
        };
        xhr.send('songId=' + encodeURIComponent(songId));
    }

    // ==================================================================
    // KEYBOARD SHORTCUTS (Accessibility)
    // ==================================================================

    function bindKeyboardShortcuts() {
        document.addEventListener('keydown', function(e) {
            // Don't intercept when typing in inputs
            if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.tagName === 'SELECT') {
                return;
            }
            switch (e.code) {
                case 'Space':
                    e.preventDefault();
                    togglePlay();
                    break;
                case 'ArrowRight':
                    if (e.shiftKey) {
                        e.preventDefault();
                        nextSong();
                    }
                    break;
                case 'ArrowLeft':
                    if (e.shiftKey) {
                        e.preventDefault();
                        prevSong();
                    }
                    break;
                case 'KeyM':
                    toggleMute();
                    break;
            }
        });
    }

    // ==================================================================
    // MODAL SYSTEM
    // ==================================================================

    window.openModal = function(modalId) {
        var modal = document.getElementById(modalId);
        if (modal) {
            modal.classList.add('active');
            var firstInput = modal.querySelector('input, select, textarea');
            if (firstInput) {
                setTimeout(function() { firstInput.focus(); }, 100);
            }
            document.body.style.overflow = 'hidden';
        }
    };

    window.closeModal = function(modalId) {
        var modal = document.getElementById(modalId);
        if (modal) {
            modal.classList.remove('active');
            document.body.style.overflow = '';
        }
    };

    // Close modal on overlay click
    document.addEventListener('click', function(e) {
        if (e.target.classList && e.target.classList.contains('modal-overlay')) {
            e.target.classList.remove('active');
            document.body.style.overflow = '';
        }
    });

    // Close modal on Escape
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') {
            document.querySelectorAll('.modal-overlay.active').forEach(function(m) {
                m.classList.remove('active');
            });
            document.body.style.overflow = '';
        }
    });

    // ==================================================================
    // CRUD: Song Management (AJAX)
    // ==================================================================

    window.editSong = function(songId) {
        // Populate modal with song data from card attributes
        var card = document.querySelector('.song-card[data-song-id="' + songId + '"]') ||
                   document.querySelector('tr[data-song-id="' + songId + '"]');

        if (!card) return;

        var title = card.getAttribute('data-song-title') || '';
        var path = card.getAttribute('data-song-path') || '';
        var plays = card.getAttribute('data-song-plays') || '0';

        document.getElementById('editSongId').value = songId;
        document.getElementById('editSongName').value = title;
        document.getElementById('editSongPath').value = path;
        document.getElementById('editSongPlays').value = plays;

        openModal('editModal');
    };

    window.deleteSong = function(songId) {
        var card = document.querySelector('.song-card[data-song-id="' + songId + '"]');
        var title = card ? card.getAttribute('data-song-title') : 'this song';

        document.getElementById('deleteSongId').value = songId;
        document.getElementById('deleteSongName').textContent = title;
        openModal('deleteModal');
    };

    // ==================================================================
    // TOAST NOTIFICATIONS
    // ==================================================================

    function showToast(message, type) {
        type = type || 'info';
        var container = document.querySelector('.toast-container');
        if (!container) {
            container = document.createElement('div');
            container.className = 'toast-container';
            document.body.appendChild(container);
        }

        var toast = document.createElement('div');
        toast.className = 'toast ' + type;

        var icons = {
            success: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16"><path d="M20 6L9 17l-5-5"/></svg>',
            error: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>',
            info: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="16" height="16"><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>'
        };

        toast.innerHTML = icons[type] + '<span>' + message + '</span>';
        container.appendChild(toast);

        setTimeout(function() {
            toast.style.animation = 'slideIn 0.3s ease-out reverse';
            setTimeout(function() { toast.remove(); }, 300);
        }, 3000);
    }

    // ==================================================================
    // UTILITIES
    // ==================================================================

    function formatTime(seconds) {
        if (isNaN(seconds) || seconds === 0) return '0:00';
        var mins = Math.floor(seconds / 60);
        var secs = Math.floor(seconds % 60);
        return mins + ':' + (secs < 10 ? '0' : '') + secs;
    }

    // ==================================================================
    // AUTO-INIT ON DOM READY
    // ==================================================================

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }

})();
