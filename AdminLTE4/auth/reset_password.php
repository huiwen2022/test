<?php
session_start();

$email = $_POST['email'] ?? '';
$new_password = $_POST['new_password'] ?? '';

$dataFile = '../users.json';
$users = file_exists($dataFile) ? json_decode(file_get_contents($dataFile), true) : [];

$found = false;
foreach ($users as &$user) {
    if ($user['email'] === $email) {
        $user['password'] = md5($new_password);
        $found = true;
        break;
    }
}
unset($user);

if ($found) {
    file_put_contents($dataFile, json_encode($users, JSON_PRETTY_PRINT));
    $_SESSION['forgot_msg'] = "密碼已更新，請使用新密碼登入。";
} else {
    $_SESSION['forgot_msg'] = "查無此 Email 註冊記錄。";
}
header("Location: ../forgot_password.php");
exit;
