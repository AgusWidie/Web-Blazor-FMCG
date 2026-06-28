// Partikel halus untuk background login
window.initLoginParticles = function () {
    const canvas = document.getElementById('loginParticles');
    if (!canvas) return;

    const ctx = canvas.getContext('2d');
    let particles = [];
    let mouseX = -1000, mouseY = -1000;
    let w, h;

    function resize() {
        w = canvas.width = window.innerWidth;
        h = canvas.height = window.innerHeight;
    }
    resize();
    window.addEventListener('resize', resize);
    document.addEventListener('mousemove', function (e) {
        mouseX = e.clientX;
        mouseY = e.clientY;
    });

    class Dot {
        constructor() { this.init(); }
        init() {
            this.x = Math.random() * w;
            this.y = Math.random() * h;
            this.r = Math.random() * 1.2 + 0.4;
            this.vx = (Math.random() - 0.5) * 0.25;
            this.vy = (Math.random() - 0.5) * 0.25;
            this.alpha = Math.random() * 0.15 + 0.05;
        }
        update() {
            var dx = this.x - mouseX;
            var dy = this.y - mouseY;
            var dist = Math.sqrt(dx * dx + dy * dy);
            if (dist < 100 && dist > 0) {
                var f = (100 - dist) / 100 * 0.4;
                this.x += (dx / dist) * f;
                this.y += (dy / dist) * f;
            }
            this.x += this.vx;
            this.y += this.vy;
            if (this.x < -20 || this.x > w + 20 || this.y < -20 || this.y > h + 20) this.init();
        }
        draw() {
            ctx.beginPath();
            ctx.arc(this.x, this.y, Math.max(0.1, this.r), 0, Math.PI * 2);
            ctx.fillStyle = 'rgba(13, 148, 136, ' + this.alpha + ')';
            ctx.fill();
        }
    }

    var count = Math.min(55, Math.floor(w * h / 22000));
    for (var i = 0; i < count; i++) particles.push(new Dot());

    function lines() {
        for (var i = 0; i < particles.length; i++) {
            for (var j = i + 1; j < particles.length; j++) {
                var dx = particles[i].x - particles[j].x;
                var dy = particles[i].y - particles[j].y;
                var d = Math.sqrt(dx * dx + dy * dy);
                if (d < 110) {
                    ctx.beginPath();
                    ctx.moveTo(particles[i].x, particles[i].y);
                    ctx.lineTo(particles[j].x, particles[j].y);
                    ctx.strokeStyle = 'rgba(13, 148, 136, ' + ((1 - d / 110) * 0.05) + ')';
                    ctx.lineWidth = 0.5;
                    ctx.stroke();
                }
            }
        }
    }

    function loop() {
        ctx.clearRect(0, 0, w, h);
        particles.forEach(function (p) { p.update(); p.draw(); });
        lines();
        requestAnimationFrame(loop);
    }
    loop();
};

// Shake animasi untuk card saat error
window.shakeLoginCard = function (id) {
    var el = document.getElementById(id);
    if (!el) return;
    el.style.animation = 'none';
    el.offsetHeight; // reflow
    el.style.animation = 'shakeCard 0.45s ease';
};