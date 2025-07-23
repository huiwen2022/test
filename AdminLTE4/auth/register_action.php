<?php
session_start();

$email = $_POST['email'] ?? '';
$password = $_POST['password'] ?? '';

if (!filter_var($email, FILTER_VALIDATE_EMAIL)) {
    $_SESSION['reg_error'] = "Invalid email format.";
    header("Location: ../register.php");
    exit;
}

$dataFile = '../users.json';
$users = file_exists($dataFile) ? json_decode(file_get_contents($dataFile), true) : [];

foreach ($users as $user) {
    if ($user['email'] === $email) {
        $_SESSION['reg_error'] = "Email already registered.";
        header("Location: ../register.php");
        exit;
    }
}

$users[] = [
    'email' => $email,
    'password' => md5($password),
];

file_put_contents($dataFile, json_encode($users, JSON_PRETTY_PRINT));
$_SESSION['user_email'] = $email;
header("Location: ../dashboard.php");
exit;
