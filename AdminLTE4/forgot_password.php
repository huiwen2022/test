<?php session_start(); ?>
<!DOCTYPE html>
<html lang="zh">
<head>
  <meta charset="UTF-8">
  <title>忘記密碼</title>
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.7/dist/css/bootstrap.min.css">
</head>
<body class="login-page bg-body-secondary">
  <div class="login-box">
    <div class="card card-outline card-warning">
      <div class="card-header text-center">
        <h1><b>忘記密碼</b></h1>
      </div>
      <div class="card-body">
        <?php if (isset($_SESSION['forgot_msg'])): ?>
          <div class="alert alert-info"><?= $_SESSION['forgot_msg']; unset($_SESSION['forgot_msg']); ?></div>
        <?php endif; ?>
        <form action="auth/reset_password.php" method="post">
          <div class="mb-3">
            <input type="email" name="email" class="form-control" placeholder="輸入註冊 Email" required>
          </div>
          <div class="mb-3">
            <input type="password" name="new_password" class="form-control" placeholder="新密碼" required>
          </div>
          <button class="btn btn-warning w-100">重設密碼</button>
        </form>
        <p class="mt-3 text-center"><a href="index.php">回到登入</a></p>
      </div>
    </div>
  </div>
</body>
</html>
