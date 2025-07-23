<?php
$page_title = "個人資料";

ob_start(); ?>
  <h1>個人資料</h1>
  <ul>
    <li>帳號 Email：<?= htmlspecialchars($_SESSION['user_email']) ?></li>
  </ul>
<?php
$content = ob_get_clean();

include 'layout/main.php';
