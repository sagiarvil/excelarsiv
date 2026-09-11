# Desteklenen PHP Framework Uzmanlık Rehberleri

Bu doküman modern PHP framework mimarilerini ve kurum içi mini-MVC kalıplarını içerir.

---

## 🚀 Laravel Framework Rehberi

# Laravel — Framework Notları

## Yapı & konvansiyon
- `app/Http/Controllers`, `app/Models`, `routes/{web,api}.php`, `database/migrations`, `resources/views` (Blade). Konvansiyonu dövme.
- Controller ince: iş mantığı → Action/Service sınıfı veya Model. Route model binding kullan.
- İsimlendirme: Model tekil (`User`), tablo çoğul (`users`), controller `UserController`, migration `create_users_table`.

## HTTP
- Doğrulama: **Form Request** sınıfı (`php artisan make:request`), controller'da `$request->validated()`. Controller'da elle `Validator::make` yapma.
- Yanıt: API için **API Resource** (`JsonResource`) — modeli ham dönme. Sayfalama `->paginate()`.
- Middleware ile kesişen kaygı (auth, throttle, rol). `throttle:60,1` rate limit.
- CSRF web'de otomatik; API'de Sanctum/Passport token.

## Eloquent
- N+1: `with(['relation'])` eager load; `->load()` sonradan. `Model::query()` üzerinden.
- Mass assignment: `$fillable` (whitelist) — `$guarded=[]` YASAK.
- Scope (`scopeActive`), accessor/mutator (`Attribute::make`), cast (`$casts` — enum, datetime, json, encrypted).
- Migration + `php artisan migrate`; geri alma için `down()`. `php artisan migrate:status`.
- Ağır sorguda `chunk`/`lazy`/`cursor`. `DB::transaction(fn () => ...)`.

## Güvenlik
- `Hash::make`/`Hash::check` (bcrypt/argon2). Policy + Gate ile yetki; `$this->authorize()`. IDOR: policy'de sahiplik.
- Blade `{{ }}` otomatik kaçış; `{!! !!}` sadece güvenilir. SQL: query builder / binding — `DB::raw` dikkat.
- `.env` repoda değil; `config()` üzerinden oku, kodda `env()` sadece config dosyasında.
- Kütle atama, `firstOrCreate`/`updateOrCreate` yarış durumu → unique index + `insertOrIgnore`.

## Performans / araç
- `config:cache`, `route:cache`, `view:cache`, `event:cache` (prod). OPcache.
- Cache (`Cache::remember`), Queue (`ShouldQueue` job) uzun işler. Horizon/telescope.
- Test: Pest/PHPUnit + `RefreshDatabase`, `Http::fake()`, factory + seeder.

## Sık hata
- `env()`'i config dışında çağırmak (config cache sonrası null).
- Eager load unutup Blade döngüsünde N+1.
- Migration'ı prod'da `migrate:fresh` (veri siler!).
- Route cache'te closure route (cache'lenemez).

---

## 🚀 Symfony Framework Rehberi

# Symfony — Framework Notları

## Yapı
- `src/Controller`, `src/Entity`, `src/Repository`, `config/`, `templates/` (Twig). Bundle'lar. Flex recipe.
- Controller ince; iş mantığı Service'te. Constructor injection (autowiring). `#[Route]` attribute (YAML değil).
- İsimlendirme PSR-4; servis id = FQCN.

## HTTP / doğrulama
- Route: `#[Route('/x/{id}', methods: ['GET'])]`. Param converter / `#[MapEntity]`.
- Doğrulama: Constraint attribute'ları Entity/DTO üstünde + `ValidatorInterface` veya Form component. Controller'da elle değil.
- Yanıt: API Platform veya `#[MapRequestPayload]` DTO + `JsonResponse`/Serializer. Pagination (KnpPaginator / Doctrine).
- CSRF: Form component token; API'de stateless + Lexik JWT.
- Güvenlik: `security.yaml` firewall + access_control + Voter (yetki, IDOR sahiplik). `#[IsGranted]`.

## Doctrine ORM
- Entity + Repository. `#[ORM\Column]` tip belirt. İlişkilerde `fetch: EXTRA_LAZY` liste sayımı için.
- N+1: DQL `JOIN ... addSelect` veya `fetch: EAGER` seçici. `->createQueryBuilder()`.
- Migration: `doctrine/migrations` — `make:migration`, `doctrine:migrations:migrate`. `up()`/`down()`.
- `EntityManager::flush()` toplu; `clear()` uzun döngüde. `wrapInTransaction()`.
- `#[ORM\Index]`, unique constraint.

## Konfig / ortam
- `.env` + `.env.local` (repoda değil). `%env(...)%` parametreleri. Secrets vault (`secrets:set`).
- Ortam: `APP_ENV` (dev/prod/test). `cache:clear`, `cache:warmup` prod deploy'da.

## Performans / araç
- OPcache + APCu (metadata). Preloading (`config/preload.php`).
- Messenger component (async, queue transport). HttpClient (retry, timeout).
- Test: PHPUnit + `WebTestCase`/`KernelTestCase`, `zenstruck/foundry` factory, `dama/doctrine-test-bundle` (transaction rollback).

## Sık hata
- Autowiring için tip yerine skaler parametre → `services.yaml`'da bind/arguments.
- Doctrine identity map: aynı entity iki kez fetch → flush çakışması.
- `dev` profiler'ı prod'a sızdırmak (debug bundle prod'da yok).
- `flush()` döngü içinde her iterasyonda (yavaş) — sonda bir kez veya batch.

---

## 🚀 CodeIgniter 4 Framework Rehberi

# CodeIgniter 4 — Framework Notları

## Yapı
- `app/Controllers`, `app/Models`, `app/Config`, `app/Views`, `app/Database/Migrations`. PSR-4 (`App\` namespace).
- Controller `BaseController` extend. `$this->request`, `$this->response`. İş mantığı Model/Service/Entity'de.
- Route: `app/Config/Routes.php` — `$routes->get('x/(:num)', 'Controller::method/$1')` veya resource/auto-routing (auto-routing-improved tercih, legacy kapat).

## Model / DB
- `CodeIgniter\Model`: `$table`, `$primaryKey`, `$allowedFields` (mass assignment whitelist — zorunlu), `$useTimestamps`, `$returnType` (Entity sınıfı önerilir).
- Query Builder: `$this->db->table('x')->where(...)->get()` — daima builder/binding, ham SQL'de `$this->db->query($sql, [$bind])`.
- Doğrulama: Model `$validationRules` veya `$this->validate($rules)` controller'da; kural dosyası `app/Config/Validation.php`.
- Migration: `php spark migrate` / `migrate:rollback`. `php spark make:migration`.
- İlişki yok (ORM light) → manuel join veya ayrı sorgu; N+1'e dikkat, `whereIn` ile topla.

## Güvenlik
- CSRF: `Config\Security` — `csrfProtection = 'session'`, form'da `csrf_field()`. AJAX'ta header.
- XSS: `esc($data)` view'da (context: `html`, `attr`, `js`, `url`). `\Config\Filters` global.
- Şifre: `password_hash`/`password_verify`. Auth için Shield (resmi).
- `.env` repoda değil; `env()` / `getenv()`. `app/Config` sınıfları ortam override.
- SQL injection: builder + binding. `$this->db->escape()` gerekirse.

## Performans / araç
- `spark optimize` (config/locator cache), OPcache. Cache: `cache()->remember()`.
- `spark serve` dev; prod'da gerçek web sunucusu, `public/` document root.
- Test: PHPUnit + `CIUnitTestCase`, `DatabaseTestTrait` (migrate + seed + transaction), `FeatureTestTrait` (HTTP).

## Sık hata
- `$allowedFields` boş/eksik → `insert`/`update` sessizce alan atlar.
- Auto-routing (legacy) açık bırakmak → beklenmedik endpoint (güvenlik).
- Entity `$casts` / `$dates` unutmak.
- `public/` dışını document root yapmak (tüm `app/` ifşa).

---

## 🚀 Yii 2 Framework Rehberi

# Yii 2 — Framework Notları

## Yapı
- `controllers/`, `models/`, `views/`, `migrations/`, `config/{web,console,db}.php`. Basic vs Advanced template.
- Controller `yii\web\Controller` (veya `yii\rest\ActiveController` API). Action → `actionX()`. İş mantığı Model/Service/Component.
- Gii ile iskele üretimi. DI container (`Yii::$container`), application component'ler config'de.

## Model / ActiveRecord
- `ActiveRecord`: `rules()` (doğrulama — zorunlu, mass assignment `safe` attribute ile kontrol), `scenarios()`, `attributeLabels()`.
- `$model->load($data)` + `$model->validate()` + `$model->save()`. `load` sadece `safe` alanları atar.
- N+1: `->with(['relation'])` eager; `joinWith` filtre için. `asArray()` salt okuma perf.
- Migration: `./yii migrate` / `migrate/down`. `safeUp()`/`safeDown()` (transaction'lı).
- Query Builder: `(new Query())->from(...)->where([...])` — dizi koşulu otomatik parametreli. Ham için `->where('x=:id',[':id'=>$id])`.

## Güvenlik
- CSRF: `enableCsrfValidation` (varsayılan açık), `Html::beginForm` token ekler. `yii\filters\VerbFilter` (HTTP method), `AccessControl` (rol/kimlik).
- RBAC (`authManager`) yetki; IDOR: `findModel()` içinde sahiplik/`checkAccess`.
- XSS: `Html::encode()` view'da. `yii\helpers\HtmlPurifier` zengin metin.
- Şifre: `Yii::$app->security->generatePasswordHash` / `validatePassword`. `security->generateRandomString` token.
- SQL: parametreli; `\yii\db\Expression` dikkat.
- Params: hassas config `params-local.php` / env, repoda değil.

## Performans / araç
- `YII_DEBUG=false`, `YII_ENV=prod`. Schema cache (`enableSchemaCache`), query cache (`$query->cache()`), fragment/page cache.
- OPcache. Asset bundle publish + minify.
- Test: Codeception (unit/functional/acceptance) veya PHPUnit; fixture'lar.

## Sık hata
- `rules()`'ta alan `safe` değil → `load()` atlar, sessiz.
- `with()` yerine döngüde ilişki erişimi → N+1.
- `YII_DEBUG` prod'da açık (bilgi ifşası, yavaş).
- `scenarios()` yanlış → beklenmedik alan doğrulaması/ataması.

---

## 🚀 Slim 4 Mikro-Framework Rehberi

# Slim 4 — Framework Notları (mikro-framework)

## Yapı
- Minimal: PSR-7 (mesaj), PSR-15 (middleware), PSR-11 (container — PHP-DI genelde). Yapıyı sen kurarsın.
- Önerilen: `src/Application/{Actions,Middleware,Handlers}`, `src/Domain`, `src/Infrastructure`, `config/`, `public/index.php`.
- Route → **Action sınıfı** (`__invoke(Request,Response,array $args)`), closure değil. Tek eylem tek sınıf (ADR deseni).

## HTTP
- `$app->get('/x/{id}', XAction::class)`. Route group + middleware.
- Request: `$request->getParsedBody()`, `getQueryParams()`, `getAttribute('id')` (route arg). Body parsing middleware ekle (JSON).
- Response: immutable — `$response->getBody()->write(...)`; `return $response->withHeader('Content-Type','application/json')->withStatus(201)`. Yardımcı responder sınıfı yaz.
- Doğrulama: kendi katmanın (respect/validation, symfony/validator, veya DTO + assert). Framework vermez.
- Hata: custom `ErrorHandler` (`errorMiddleware`), PSR-3 logger, prod'da detay sızdırma yok, `ProblemDetails` (RFC 7807) uygula.

## DI / servisler
- Container tanımları `config/container.php`. Constructor injection. Autowire (PHP-DI) + explicit factory.
- DB: PDO/Doctrine DBAL/Cycle/Eloquent'i sen bağlarsın (illuminate/database standalone da olur).

## Güvenlik
- CSRF middleware (slim/csrf) form uygulamalarında; API'de token auth middleware.
- CORS middleware explicit (tuura/cors) — origin beyaz listesi.
- Header güvenliği middleware (CSP, HSTS...). Rate limit middleware.
- SQL: seçtiğin katmanda prepared statement. XSS: şablon motorunda (Twig) autoescape.

## Performans / araç
- FastRoute route cache (`routeCollector->setCacheFile`). OPcache. Container compile (PHP-DI `enableCompilation`).
- Test: PHPUnit — `$app->handle($request)` ile action testi; slim/psr7 ile Request kur.

## Sık hata
- Response immutable — `$response->withStatus(201)` sonucunu return etmeyip eskisini dönmek.
- Body parsing middleware eklemeden `getParsedBody()` null.
- Middleware sırası (LIFO ekleme) — auth/CORS/hata sırasını yanlış kurmak.
- Her şeyi closure route'a yığmak (test edilemez, cache'lenemez).

---

## 🚀 WordPress Tema & Eklenti Rehberi

# WordPress (tema/eklenti) — Framework Notları

## Yapı & çalışma
- Çekirdeği değiştirme. Özelleştirme: **eklenti** (işlev) veya **child theme** (görünüm). `functions.php`, `plugins/<slug>/`.
- Hook sistemi: `add_action('init', ...)`, `add_filter('the_content', ...)`. Çekirdek/başka eklentiyi hook ile genişlet.
- Enqueue: `wp_enqueue_script/style` (`wp_enqueue_scripts` hook) — asla elle `<script>`. Sürüm + bağımlılık ver.
- Ayarlar: Settings API (`register_setting`, `add_settings_field`), Options API (`get_option`/`update_option`), transient (cache) `get_transient`.

## Veri / DB
- `WP_Query` / `get_posts` (ana döngü için `pre_get_posts` filtresi). `WP_User_Query`, `WP_Term_Query`.
- Özel sorgu: `$wpdb->prepare("... WHERE id=%d", $id)` — **daima prepare**, `$wpdb->get_results/get_var`.
- Meta: `get_post_meta`/`update_post_meta` (autoload dikkat). Çok/ağır veri için özel tablo (`dbDelta` ile şema).
- `WP_Query` `meta_query`/`tax_query` pahalı — indeks yok; ölçekte özel tablo/ETL.

## Güvenlik (WP'ye özgü)
- **Nonce**: her form/ajax/işlem → `wp_nonce_field()` + `check_admin_referer()` / `wp_verify_nonce()`. (WP'nin CSRF'i budur.)
- **Yetki**: `current_user_can('edit_post', $id)` her yazma/silmede (IDOR).
- **Kaçış (çıktıda)**: `esc_html`, `esc_attr`, `esc_url`, `esc_js`, `wp_kses_post` (zengin metin whitelist). Çeviri + kaçış: `esc_html__()`.
- **Temizleme (girdide)**: `sanitize_text_field`, `sanitize_email`, `absint`, `wp_kses`.
- AJAX: `wp_ajax_{action}` / `wp_ajax_nopriv_{action}`; nonce + yetki içeride. REST: `permission_callback` zorunlu (asla `__return_true` hassas uçta).
- `ABSPATH` guard dosya başında. Doğrudan dosya erişimini engelle.
- SQL: `$wpdb->prepare`; `esc_sql` sadece son çare.

## Performans / araç
- Object cache (Redis/Memcached) + transient. Sorgu sayısı (Query Monitor). Autoload'lu şişkin option temizliği.
- Sayfa cache (eklenti/edge). `WP_DEBUG` prod'da kapalı, `WP_DEBUG_LOG` ile logla.
- `wp-cli` (`wp db`, `wp cron`, `wp plugin`). Gerçek cron (`DISABLE_WP_CRON` + sistem cron) yüksek trafikte.
- Test: `WP_Mock` / Brain Monkey (birim), `wp-phpunit` (entegrasyon).

## Sık hata
- Nonce/yetki kontrolü olmadan işlem (CSRF/IDOR).
- Çıktı kaçışsız (`echo $meta`).
- `functions.php`'de ağır kod (her istekte). Hook'a bağla, koşullu çalıştır.
- `pre_get_posts`'ta `is_admin()`/`is_main_query()` kontrolü unutup admin sorgularını bozmak.
- Çekirdek/tema dosyasını doğrudan düzenlemek (güncellemede kaybolur).

---

## 🚀 Beta Framework (Özel Mini-MVC) Rehberi

# Beta Framework (kurum içi / özel PHP mini-MVC) — Notlar

Bizim framework. Saf PHP, harici framework yok. Bilinen kalıplar (proje kodundan; yeni detay çıkarsa bu dosyayı güncelle).

## Dizin
- `app/Core/` — çekirdek: `Controller`, `Database`, `Request`, `Response`, `Router`, `Auth`, `CustomerAuth`, `Totp` vb.
- `app/Controllers/{Admin,Customer,Api,...}/` — controller'lar (namespace `App\Controllers\...`).
- `app/Models/` — model'ler (namespace `App\Models`), ham PDO.
- `app/routes.php` — rota tablosu.
- `app/Helpers/functions.php` — global yardımcılar (`flash`, `config`, `e`, `ek_lisans_kademeleri` ...).
- `config/` — `app.php`, `database.php`.
- `database/Migrations/NNN_ad.sql`, `database/Seeds/`, `database/migrate.php`, `database/seed.php`.
- `themes/{Admin,Site}/default/` — görünümler: `layout/main.php`, `partials/`, `<bölüm>/<view>.php`.

## Rota
- `app/routes.php` bir dizi döndürür; her satır: `['GET'|'POST', '/yol/{param}', [Controller::class, 'method']]`.
- Router: `/admin` veya `/admin/` ile başlayan yollar (`/admin/login` hariç) `Auth::requireLogin()` çağırır; `/client/*` → `CustomerAuth::requireLogin()`; `/api/*` → ApiKeyGuard. Ek middleware yazma.
- Literal rota aynı METHOD'ta `{param}`'lı olandan önce/net tanımlanmalı (ör. `/admin/x/bulk-delete` POST ile `/admin/x/{id}` GET çakışmaz).

## Controller
- `class X extends \App\Core\Controller`.
- `$this->view('dir/file', [...])` → Admin/Site layout ile render (`Response::html`).
- `$this->viewClient('client/file', [...])` → `Response::render('Site/default', ...)`.
- `$this->redirect('/yol')` / `Response::redirect(...)`, `$this->json($data, $status)`.
- `$this->csrfValidate($request)` — yoksa `Auth::verifyCsrf($request->input('_csrf'))`.
- `flash('success'|'danger', 'mesaj')` → sonraki istekte gösterilir; sonra mutlaka `redirect`.

## Request
- `$request->param('id')` (rota parametresi), `$request->post('k', $default)`, `$request->input('k')` (post+get+json), `$request->get('k')` (query), `$request->files('name')`.

## Model / DB
- `namespace App\Models; use App\Core\Database;`
- `Database::pdo()->prepare('SELECT ... WHERE id = ?')->execute([$id])` — daima prepared statement. `SELECT *` yerine kolon say.
- Fetch: PDO varsayılan; `->fetch()` / `->fetchAll()`.
- İlişki/ORM yok — JOIN veya `WHERE id IN (...)` ile N+1 önle.
- **Migration çift adım:** `database/Migrations/NNN_ad.sql` yaz **ve** aktif DB'ye uygula (PDO ile, `config/database.php`). PHP CLI: `php`.
- ENUM/kolon: kodun beklediği değerlerle şema birebir aynı olmalı. Model `create()`/`update()` alan listesi tabloda var olan kolonlarla eşleşmeli (yoksa PDOException).

## Görünüm
- PHP şablon; `<?= htmlspecialchars($x) ?>` / `e($x)`. `Auth::csrfToken()` form'da hidden `_csrf`.
- Tailwind (CDN) + Alpine.js + Bootstrap Icons. Açık tema. Admin tablo standardı → `refs/frontend-developer.md`.

## Auth
- Admin: `App\Core\Auth` — `Auth::check()`, `Auth::id()`, `Auth::username()`, `Auth::csrfToken()`, `Auth::verifyCsrf()`, `Auth::requireLogin()`. Oturum: `$_SESSION['admin_*']`. `admins` tablosu.
- Müşteri: `App\Core\CustomerAuth` — `id()`, `name()`, `check()`, `requireLogin()`. `users` tablosu. `AdminCustomerController::impersonate` ile admin müşteri oturumu açabilir.

## Config
- `config('app.x')`, `config('database.*')`. Hassas değer config dosyasında/env; kodda gömme, repoya sokma.

## Doğrulama (bu framework'te)
1. `php -l <dosya>`.
2. Rota kayıtlı + controller method `public function` + isim birebir.
3. Migration yazıldıysa aktif DB'ye uygulandı; enum/kolon uyumu.
4. `flash` sonrası `redirect` var mı.
5. Tarayıcıda ilgili aksiyonu çalıştır, flash + sonuç + konsol.

## Sık hata (bu framework)
- Migration'ı sadece dosyaya yazıp DB'ye uygulamamak.
- `Response`/`Auth` sınıfını `use` etmeden kullanmak.
- Rota param adı ↔ `$request->param('...')` farkı.
- View'a değişken geçmeyi unutup şablonda undefined.
