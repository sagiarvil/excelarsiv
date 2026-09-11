# Veritabanı Modülü — ag-php-developer

Bu modül `ag-php-developer` şemsiyesinin bir parçasıdır.
Derin DBA görevleri için `ag-database-expert` skill'ini kullan.

---

## 1. PHP-MySQL Entegrasyon Standartları

### PDO Bağlantısı

```php
<?php
declare(strict_types=1);

$pdo = new PDO(
    'mysql:host=localhost;dbname=mydb;charset=utf8mb4',
    $_ENV['DB_USER'],
    $_ENV['DB_PASS'],
    [
        PDO::ATTR_ERRMODE            => PDO::ERRMODE_EXCEPTION,
        PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
        PDO::ATTR_EMULATE_PREPARES   => false,
    ]
);
```

### Prepared Statement Zorunluluğu

```php
// ✅ DOĞRU
$stmt = $pdo->prepare('SELECT * FROM users WHERE email = :email');
$stmt->execute([':email' => $email]);
$user = $stmt->fetch();

// ❌ YASAK — SQL Injection
$result = $pdo->query("SELECT * FROM users WHERE email = '$email'");
```

---

## 2. Migration Kuralı (Çift Adımlı)

```
1. database/Migrations/NNN_ad.sql dosyasını oluştur (bir sonraki numara)
2. Aktif veritabanına doğrudan PDO ile uygula
3. Geri alma (rollback) SQL'ini de yaz
```

```sql
-- database/Migrations/001_create_users.sql
CREATE TABLE users (
    id         BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    email      VARCHAR(255) NOT NULL UNIQUE,
    created_at TIMESTAMP    NOT NULL DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
```

---

## 3. N+1 Sorgu Yasağı

```php
// ❌ YASAK — N+1
foreach ($orders as $order) {
    $user = $pdo->query("SELECT * FROM users WHERE id = {$order['user_id']}")->fetch();
}

// ✅ DOĞRU — JOIN ile tek sorgu
$stmt = $pdo->prepare('
    SELECT o.*, u.name, u.email
    FROM orders o
    JOIN users u ON u.id = o.user_id
    WHERE o.status = :status
');
$stmt->execute([':status' => 'active']);
$orders = $stmt->fetchAll();
```

---

## 4. Sayfalama — Keyset (Cursor)

```php
// ❌ YASAK — büyük tabloda yavaş
$stmt = $pdo->prepare('SELECT * FROM posts LIMIT :limit OFFSET :offset');

// ✅ DOĞRU — keyset pagination
$stmt = $pdo->prepare('
    SELECT * FROM posts
    WHERE id > :last_id
    ORDER BY id ASC
    LIMIT :limit
');
$stmt->execute([':last_id' => $lastId, ':limit' => $perPage]);
```

---

## 5. Veritabanı Kolonu Tipleri (Best Practice)

| Tip | Kullanım |
|---|---|
| `BIGINT UNSIGNED` | Primary key, foreign key |
| `VARCHAR(255)` | E-posta, isim, slug |
| `TEXT` | Uzun metin (≤65KB) |
| `MEDIUMTEXT` | Makale içeriği |
| `DECIMAL(10,2)` | Para birimi (FLOAT değil!) |
| `TINYINT(1)` | Boolean flag |
| `TIMESTAMP` | created_at, updated_at (UTC) |
| `JSON` | MySQL 5.7+ yapılandırma/metadata |

---

## 6. İndeks Stratejisi

```sql
-- Composite index: sık birlikte sorgulanan kolonlar
CREATE INDEX idx_orders_user_status ON orders (user_id, status);

-- Covering index: SELECT kolonları da index'e dahil
CREATE INDEX idx_posts_slug_title ON posts (slug, title);

-- EXPLAIN ile doğrula
EXPLAIN SELECT * FROM orders WHERE user_id = 1 AND status = 'active';
```
