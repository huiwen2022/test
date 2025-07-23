<?php
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
?>
<!-- 下方為登入後畫面 -->
<!DOCTYPE html>
<html lang="zh">
<head>
  <meta charset="UTF-8">
  <title>控制台</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <!-- AdminLTE & Bootstrap -->
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.7/dist/css/bootstrap.min.css">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/admin-lte@4.0.0-rc3/dist/css/adminlte.min.css">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.13.1/font/bootstrap-icons.css">
</head>
<body class="hold-transition sidebar-mini">
<div class="wrapper">
<!-- Navbar -->
<nav class="main-header navbar navbar-expand navbar-white navbar-light">
  <ul class="navbar-nav">
    <li class="nav-item"><a class="nav-link" href="#">🏠 控制台</a></li>
  </ul>
  <ul class="navbar-nav ms-auto">
    <li class="nav-item">
      <a class="nav-link" href="logout.php">登出</a>
    </li>
  </ul>
</nav>