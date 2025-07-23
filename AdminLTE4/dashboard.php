<?php
$page_title = "控制台";

ob_start(); ?>
  <h1>控制台首頁</h1>
  <p>歡迎回來，<?= htmlspecialchars($_SESSION['user_email']) ?>！</p>
<?php
$content = ob_get_clean();

include 'layout/main.php';
