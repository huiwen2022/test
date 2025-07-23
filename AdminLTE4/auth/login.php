<?php
ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);
error_reporting(E_ALL);

session_start();

$email = $_POST['email'] ?? '';
$password = $_POST['password'] ?? '';

$dataFile = '../users.json';
$users = file_exists($dataFile) ? json_decode(file_get_contents($dataFile), true) : [];

$found = false;
foreach ($users as $user) {
    if ($user['email'] === $email && $user['password'] === md5($password)) {
        $_SESSION['user_email'] = $email;
        $found = true;
        break;
    }
}

if ($found) {
    header("Location: ../dashboard.php");
} else {
    $_SESSION['error'] = "Invalid credentials.";
    header("Location: ../index.php");
}
exit;
