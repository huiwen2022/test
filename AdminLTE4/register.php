<?php session_start(); ?>
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Register</title>
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.7/dist/css/bootstrap.min.css">
</head>
<body class="container mt-5">
  <h2>Register</h2>
  <?php if (isset($_SESSION['reg_error'])): ?>
    <div class="alert alert-danger"><?= $_SESSION['reg_error']; unset($_SESSION['reg_error']); ?></div>
  <?php endif; ?>
  <form action="auth/register_action.php" method="post">
    <div class="mb-3">
      <label>Email:</label>
      <input type="email" name="email" class="form-control" required>
    </div>
    <div class="mb-3">
      <label>Password:</label>
      <input type="password" name="password" class="form-control" required>
    </div>
    <button class="btn btn-success">Register</button>
  </form>
  <p class="mt-3"><a href="index.php">Back to Login</a></p>
</body>
</html>
