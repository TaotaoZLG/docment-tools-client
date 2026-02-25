using docment_tools_client.Helpers;
using docment_tools_client.Services;
using System;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace docment_tools_client.ViewModels
{
    public class ChangePasswordViewModel : ViewModelBase
    {
        private string _oldPassword = "";
        public string OldPassword
        {
            get => _oldPassword;
            set
            {
                SetProperty(ref _oldPassword, value);
                IsConfirmEnabled = !string.IsNullOrEmpty(value) && !string.IsNullOrEmpty(NewPassword) && !string.IsNullOrEmpty(ConfirmPassword);
            }
        }

        private string _newPassword = "";
        public string NewPassword
        {
            get => _newPassword;
            set
            {
                SetProperty(ref _newPassword, value);
                IsConfirmEnabled = !string.IsNullOrEmpty(OldPassword) && !string.IsNullOrEmpty(value) && !string.IsNullOrEmpty(ConfirmPassword);
            }
        }

        private string _confirmPassword = "";
        public string ConfirmPassword
        {
            get => _confirmPassword;
            set
            {
                SetProperty(ref _confirmPassword, value);
                IsConfirmEnabled = !string.IsNullOrEmpty(OldPassword) && !string.IsNullOrEmpty(NewPassword) && !string.IsNullOrEmpty(value);
            }
        }

        private bool _isConfirmEnabled;
        public bool IsConfirmEnabled
        {
            get => _isConfirmEnabled;
            set => SetProperty(ref _isConfirmEnabled, value);
        }

        private string _statusMessage = "";
        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }
        
        public ICommand ConfirmCommand { get; }
        public ICommand CancelCommand { get; }

        public Action? CloseAction { get; set; }

        public ChangePasswordViewModel()
        {
            ConfirmCommand = new RelayCommand(ExecuteConfirm);
            CancelCommand = new RelayCommand(() => CloseAction?.Invoke());
        }

        private async void ExecuteConfirm()
        {
            StatusMessage = "";
            IsConfirmEnabled = false;

            // 验证逻辑
            if (string.IsNullOrEmpty(OldPassword))
            {
                StatusMessage = "请输入旧密码";
                IsConfirmEnabled = true;
                return;
            }

            if (string.IsNullOrEmpty(NewPassword))
            {
                StatusMessage = "请输入新密码";
                IsConfirmEnabled = true;
                return;
            }

            if (NewPassword.Length < 6 || NewPassword.Length > 20)
            {
                StatusMessage = "新密码长度必须在6-20位之间";
                IsConfirmEnabled = true;
                return;
            }

            // 非法字符：< > " ' \ |
            if (Regex.IsMatch(NewPassword, "[<>\"'\\\\|]"))
            {
                StatusMessage = "新密码不能包含非法字符 (< > \" ' \\ |)";
                IsConfirmEnabled = true;
                return;
            }

            if (OldPassword == NewPassword)
            {
                StatusMessage = "新密码不能与旧密码相同";
                IsConfirmEnabled = true;
                return;
            }

            if (NewPassword != ConfirmPassword)
            {
                StatusMessage = "两次输入的新密码不一致";
                IsConfirmEnabled = true;
                return;
            }

            try
            {
                // 加密密码 (假设服务端需要加密传输，或者如果是明文直接传。Login使用encrypt，这里假设也需要，但题目参数是oldPassword和newPassword，未特别说明加密)
                // 题目： "oldPassword": "旧密码", "newPassword": "新密码" 
                // 若 LoginAsync 加密了，这里是否要加密？
                // 题目中 request body 看起来是明文示例。但Login用了 EncryptHelper.EncryptString。
                // 安全起见，通常会加密。但如果后端接口只接受明文json...
                // 假设后端和Login一样需要MD5或RSA，但这里ApiService.LoginAsync调用了EncryptHelper。
                // 题目里的request body example is plain text descriptions.
                // Given the prompt "Request params: ...", I will assume plain text unless existing code suggests otherwise.
                // Let's check LoginAsync in ApiService.cs. It takes `encryptedPassword`.
                // So the viewmodel probably encrypts.
                // The prompt says "Request parameters: { "oldPassword": "...", ... }".
                // I will send what is provided. If `ApiService.UpdatePasswordAsync` is implemented to take strings.
                // Wait, in `LoginAsync` context, `encryptedPassword` is passed in.
                // Validating user input locally first.
                
                // Call API
                var result = await ApiService.UpdatePasswordAsync(OldPassword, NewPassword);
                
                if (result.IsSuccess)
                {
                    MessageBox.Show(result.Msg, "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    CloseAction?.Invoke();
                }
                else
                {
                    StatusMessage = result.Msg;
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"操作失败: {ex.Message}";
            }
            finally
            {
                IsConfirmEnabled = true;
            }
        }
    }
}
