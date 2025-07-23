<?php
if (!isset($page_title)) $page_title = '未命名頁面';
if (!isset($content)) $content = '';
session_start();
if (!isset($_SESSION['user_email'])) {
    echo '<!DOCTYPE html>
    <html lang="zh">
    <head>
        <meta charset="UTF-8">
        <title>請先登入</title>
        <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.7/dist/css/bootstrap.min.css">
    </head>
    <body class="d-flex align-items-center justify-content-center vh-100 bg-light">
      <div class="text-center">
        <h3 class="text-danger mb-3">⚠ 請先登入</h3>
        <a href="index.php" class="btn btn-primary">點選此處回到登入頁面</a>
      </div>
    </body>
    </html>';
    exit;
}
$current_page = basename($_SERVER['PHP_SELF']);
?>
<!DOCTYPE html>
<html lang="zh">
<head>
  <meta charset="UTF-8">
  <title><?= htmlspecialchars($page_title) ?> | AdminLTE JSON</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.7/dist/css/bootstrap.min.css">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/admin-lte@4.0.0-rc3/dist/css/adminlte.min.css">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.13.1/font/bootstrap-icons.css">
</head>
<body class="hold-transition sidebar-mini">
<div class="wrapper">

<!-- Navbar -->
<nav class="main-header navbar navbar-expand navbar-white navbar-light shadow-sm">
  <div class="container-fluid">
    <span class="navbar-brand"><strong><?= $page_title ?></strong></span>
    <ul class="navbar-nav ms-auto">
      <li class="nav-item">
        <span class="nav-link"><?= htmlspecialchars($_SESSION['user_email']) ?></span>
      </li>
      <li class="nav-item">
        <a class="nav-link text-danger" href="logout.php"><i class="bi bi-box-arrow-right"></i> 登出</a>
      </li>
    </ul>
  </div>
</nav>

<!-- Sidebar -->
<aside class="main-sidebar sidebar-dark-primary elevation-4">
  <a href="#" class="brand-link text-center">
    <span class="brand-text fw-bold">🧩 AdminLTE JSON</span>
  </a>
  <div class="sidebar">
    <nav class="mt-2">
      <ul class="nav nav-pills nav-sidebar flex-column" role="menu">
        <li class="nav-item">
          <a href="dashboard.php" class="nav-link <?= $current_page === 'dashboard.php' ? 'active' : '' ?>">
            <i class="nav-icon bi bi-speedometer2"></i>
            <p>控制台</p>
          </a>
        </li>
        <li class="nav-item">
          <a href="profile.php" class="nav-link <?= $current_page === 'profile.php' ? 'active' : '' ?>">
            <i class="nav-icon bi bi-person-circle"></i>
            <p>個人資料</p>
          </a>
        </li>
      </ul>
    </nav>
  </div>
</aside>

<!-- Content Wrapper -->
<div class="content-wrapper">
  <section class="content pt-4">
    <div class="container-fluid">
      <?= $content ?>
    </div>
  </section>
</div>

</div> <!-- /.wrapper -->

<!-- Scripts -->
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.7/dist/js/bootstrap.bundle.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/admin-lte@4.0.0-rc3/dist/js/adminlte.min.js"></script>
</body>
</html>
