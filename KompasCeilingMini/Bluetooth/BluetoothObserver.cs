using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Timers;
using Windows.Devices.Bluetooth;
using Windows.Devices.Bluetooth.Advertisement;
using Windows.Devices.Bluetooth.GenericAttributeProfile;

using Windows.Security.Cryptography;


namespace BlueTest.BlueClass
{
    public class BluetoothObserver
    {
        public static BluetoothLEAdvertisementWatcher Watcher { get => _watcher; set => _watcher = value; }
        private string _services;
        private string _characteristic;
        private string _macdevice;
        private string _namedevice;
        private static BluetoothLEDevice _bluetoothLeDevice;
        private static BluetoothLEAdvertisementWatcher _watcher;
        private System.Timers.Timer aTimer;

        //Получить статус

        public event EventHandler<bool> ConnectChange;

        public event EventHandler<string> statlabelChange;

        private List<string> GetAsyncList { get; set; } //Лист размеров

        //Вызов листа размеров

        //
        //Последний замер
        //
        public string lastDim { get; private set; }

        public event EventHandler<double> onLastDimention;


        //Получить последний размер


        //Список всех замеров

        public void Start(string NameDevice, string MacDevice, string Services, string Characteristic)
        {

            if (_bluetoothLeDevice == null)
            {
                _services = Services;
                _characteristic = Characteristic;
                _macdevice = MacDevice;
                _namedevice = NameDevice;

                _watcher = new BluetoothLEAdvertisementWatcher()
                {
                    ScanningMode = BluetoothLEScanningMode.Active
                };
                _watcher.SignalStrengthFilter.InRangeThresholdInDBm = -50;
                _watcher.SignalStrengthFilter.InRangeThresholdInDBm = -90;
                _watcher.SignalStrengthFilter.OutOfRangeTimeout = TimeSpan.FromSeconds(2);
                _watcher.Received += Watcher_Received;
                _watcher.Stopped += Watcher_Stopped;

                statlabelChange(this, "Запускаем");
                _watcher.Start();
            }
            else GattDevice(_bluetoothLeDevice.BluetoothAddress);
        }

        public bool isFindDevice { get; set; } = false;
        private void Watcher_Received(BluetoothLEAdvertisementWatcher sender, BluetoothLEAdvertisementReceivedEventArgs args)
        {
            /*statlabelChange(this, "Сканируем..");

            if (isFindDevice)
                return;*/
            if (args.Advertisement.LocalName.Contains(_namedevice))
                GattDevice(args.BluetoothAddress);
        }

        private async void GattDevice(ulong adress)
        {
            statlabelChange(this, "Нашли");
            _bluetoothLeDevice = await BluetoothLEDevice.FromBluetoothAddressAsync(adress);

            GC.KeepAlive(_bluetoothLeDevice);
            if (_bluetoothLeDevice.DeviceId == _macdevice)
            {
                GattDeviceServicesResult gattService = await _bluetoothLeDevice.GetGattServicesForUuidAsync(new Guid(_services), BluetoothCacheMode.Cached);

                //GattDeviceServicesResult GattServices = await _bluetoothLeDevice.GetGattServicesAsync();//Получаем сервисы

                _bluetoothLeDevice.ConnectionStatusChanged += onBluetoothConnectionStatus;

                if (gattService.Status == GattCommunicationStatus.Success)
                {
                    var service = gattService.Services.First();

                    GattCharacteristicsResult characteristicsResult = await service.GetCharacteristicsForUuidAsync(new Guid(_characteristic), BluetoothCacheMode.Cached);

                    if (characteristicsResult.Status == GattCommunicationStatus.Success)
                    {
                        var characteristic = characteristicsResult.Characteristics.First();

                        GattCharacteristicProperties properties = characteristic.CharacteristicProperties;

                        if (properties.HasFlag(GattCharacteristicProperties.Notify))
                        {
                            try
                            {
                                var notifyResult = await characteristic.WriteClientCharacteristicConfigurationDescriptorAsync(
                                      GattClientCharacteristicConfigurationDescriptorValue.Notify);
                                if (notifyResult == GattCommunicationStatus.Success)
                                {
                                    statlabelChange(this, "Подключено");
                                    characteristic.ValueChanged += Charac_ValueChangedAsync;
                                    GC.KeepAlive(characteristic);
                                    isFindDevice = true;
                                    _watcher.Stop();
                                    SetTimer();
                                    return;
                                }
                            }
                            catch { _watcher.Stop(); return; }
                        }
                    }
                }
            }
        }
            
        private void SetTimer()
        {
            aTimer = new System.Timers.Timer(TimeSpan.FromSeconds(30).TotalMilliseconds);
            aTimer.Elapsed += OnTimerEvent;
            aTimer.AutoReset = true;
            aTimer.Start();
        }

        private void OnTimerEvent(object sender, ElapsedEventArgs e)
        {
            Start(_namedevice, _macdevice, _services, _characteristic);
        }

        private void onBluetoothConnectionStatus(BluetoothLEDevice sender, object args)
        {
            if (sender.ConnectionStatus == BluetoothConnectionStatus.Disconnected)
            {
                if (aTimer != null)
                {
                    aTimer.Stop();
                    aTimer.Dispose();
                }
                if (_bluetoothLeDevice != null)
                {
                    _bluetoothLeDevice.ConnectionStatusChanged -= onBluetoothConnectionStatus;
                    _bluetoothLeDevice.Dispose();
                }
                isFindDevice = false;
                _bluetoothLeDevice = null;
                GC.Collect();
                if (ConnectChange!= null) ConnectChange(this, false);
                statlabelChange(this, "Отключено");
            }
            else if (sender.ConnectionStatus == BluetoothConnectionStatus.Connected)
            {
                statlabelChange(this, "Подключено");
                _watcher.Stop();
            }
        }

        private async void Charac_ValueChangedAsync(GattCharacteristic sender, GattValueChangedEventArgs args)
        {
            CryptographicBuffer.CopyToByteArray(args.CharacteristicValue, out byte[] data);
            string dataFromNotify;
            try
            {
                //Asuming Encoding is in ASCII, can be UTF8 or other!
                dataFromNotify = Encoding.ASCII.GetString(data);
                //Debug.Write(dataFromNotify);
            }
            catch (ArgumentException)
            {
                //Debug.Write("Unknown format");
            }
            GattReadResult dataFromRead = await sender.ReadValueAsync();
            CryptographicBuffer.CopyToByteArray(dataFromRead.Value, out byte[] dataRead);
            string dataFromReadResult;
            try
            {
                //Asuming Encoding is in ASCII, can be UTF8 or other!
                dataFromReadResult = Encoding.ASCII.GetString(dataRead);
                onLastDimention(this, double.Parse(Regex.Replace(dataFromReadResult, "[A-Za-z \n\r\0]", "").Replace('.', ',')));

                ///"DATA FROM READ: " + dataFromReadResult);
            }
            catch (ArgumentException)
            {
                //Debug.Write("Unknown format");
            }

        }


        private void Watcher_Stopped(BluetoothLEAdvertisementWatcher sender, BluetoothLEAdvertisementWatcherStoppedEventArgs args)
        {
        }

    }
}
