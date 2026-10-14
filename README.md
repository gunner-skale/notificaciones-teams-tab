# consegsa-teams-tab

Pestaña de Teams de la app **Consegsa Tareas**. Al hacer clic en un aviso de tarea en Teams, se abre esta página y
muestra el detalle de la tarea.

- Node 24, sin dependencias npm ni build.
- `server.js` sirve `public/` y reenvía `/api/tarea` al backend (`GET /integracion/teams/tab/tarea?t=`). El
  navegador solo habla con este servicio, así que no hay CORS.
- Cada aviso lleva un **enlace firmado** por el backend (`TEAMS_TAB_SECRET`). Teams se lo pasa a la página como
  `subPageId`, y la tarea se muestra con lo que ese usuario puede ver.
- **Requisito:** `BACKEND_URL` tiene que ser **el mismo backend que envía los avisos**. Si no, la firma no coincide y
  la página dice "enlace no válido".

## Desplegar en el servidor del reverse proxy
```bash
# copiar la carpeta al servidor, p. ej.:
scp -r consegsa-teams-tab adminia@192.168.11.42:~/proyectos/

# en el servidor
cd ~/proyectos/consegsa-teams-tab
cp .env.example .env        # poner BACKEND_URL
sudo docker compose up -d --build
sudo docker compose ps      # debe quedar "healthy"
curl http://127.0.0.1:6090/salud   # {"ok":true}
```

Comprobar que el contenedor llega al backend:
```bash
sudo docker compose exec teams-tab wget -qO- "http://127.0.0.1:6090/api/tarea?t=x"
# {"code":401,...}  -> llega al backend (el enlace "x" es inválido, es lo esperado)
# {"code":502,...}  -> no llega: revisar BACKEND_URL
# {"code":404,...} o 403 -> el backend no tiene todavía /integracion/teams/tab (desplegar la versión nueva)
```

## nginx (`apps.conf.template`, dentro del server de `${DOMAIN}`, antes de `location /`)
```nginx
    # ── Teams: detalle de tareas que abre la notificación (consegsa-teams-tab) ──
    location = /notifyTeams { return 301 /notifyTeams/; }

    location ^~ /notifyTeams/ {
        proxy_pass http://127.0.0.1:6090/;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        proxy_redirect off;
        proxy_read_timeout 30s;
        proxy_connect_timeout 10s;

        # Teams muestra esta página dentro de un iframe: aquí NO va el X-Frame-Options ni el CSP del server
        # (frame-ancestors 'self' la bloquearía). El CSP correcto lo envía server.js.
        add_header X-Content-Type-Options "nosniff" always;
        add_header Referrer-Policy "no-referrer" always;
        add_header Strict-Transport-Security "max-age=31536000" always;
    }
```
`127.0.0.1` sirve si el contenedor de nginx usa `network_mode: host`, como parecen indicar los demás `proxy_pass`
a `127.0.0.1`. Para saberlo: `grep -n network_mode docker-compose.yml` en la carpeta del reverse proxy.
Si nginx **no** usa la red del host:
- En `docker-compose.yml` de este servicio, publicar `"6090:6090"` (sin `127.0.0.1`) y usar
  `proxy_pass http://192.168.11.42:6090/;`, o
- unir los dos servicios a la misma red de Docker y usar `proxy_pass http://consegsa-teams-tab:6090/;`.

Después: `sudo docker compose run --rm --no-deps reverse-proxy nginx -t` y
`sudo docker compose up -d --force-recreate reverse-proxy`.

## Probar en local, sin Docker
```bash
PORT=6090 BACKEND_URL=http://localhost:6080 node server.js
# http://localhost:6090/?t=<enlace firmado>
```

## Manifiesto de Teams
En el repo del backend, `teams-app/manifest.json`:
- `staticTabs[0].contentUrl` y `websiteUrl` → `https://<dominio>/notifyTeams/`
- `validDomains` → `["<dominio>"]`

Si cambia algo, subir `version`, regenerar el zip y actualizarlo en Teams admin center.
