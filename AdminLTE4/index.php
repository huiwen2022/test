<?php session_start(); ?>
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Login</title>
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.7/dist/css/bootstrap.min.css">
</head>
<body class="container mt-5">
  <h2>Login</h2>
  <?php if (isset($_SESSION['error'])): ?>
    <div class="alert alert-danger"><?= $_SESSION['error']; unset($_SESSION['error']); ?></div>
  <?php endif; ?>
    <form action="auth/login.php" method="post">
    <div class="input-group mb-1">
        <div class="form-floating">
        <input name="email" type="email" class="form-control" id="loginEmail" placeholder="Email" required />
        <label for="loginEmail">Email</label>
        </div>
        <div class="input-group-text"><span class="bi bi-envelope"></span></div>
    </div>
    <div class="input-group mb-1">
        <div class="form-floating">
        <input name="password" type="password" class="form-control" id="loginPassword" placeholder="Password" required />
        <label for="loginPassword">Password</label>
        </div>
        <div class="input-group-text"><span class="bi bi-lock-fill"></span></div>
    </div>
    <div class="row mt-3">
        <div class="col-8 d-inline-flex align-items-center">
        <div class="form-check">
            <input class="form-check-input" type="checkbox" value="" id="remember" />
            <label class="form-check-label" for="remember"> Remember Me </label>
        </div>
        </div>
        <div class="col-4">
        <div class="d-grid gap-2">
            <button type="submit" class="btn btn-primary">Sign In</button>
        </div>
        </div>
    </div>
    </form>
    <p class="mb-1 mt-2"><a href="forgot_password.php">忘記密碼？</a></p>
    <p class="mb-0"><a href="register.php" class="text-center">註冊新帳號</a></p>
</body>
</html>
