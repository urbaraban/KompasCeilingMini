using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Windows.Devices.Bluetooth;
using Windows.Devices.Bluetooth.Advertisement;
using Windows.Devices.Bluetooth.GenericAttributeProfile;
using Windows.Security.Cryptography;


namespace BlueTest.BlueClass
{
    public class BluetoothObserver : INotifyPropertyChanged
    {
        BluetoothLEAdvertisementWatcher Watcher { get; set; }
        private string _services;
        private string _characteristic;
        private string _namedevice;
        private static BluetoothLEAdvertisementWatcher _watcher;

        //Получить статус
        private bool GetAsyncConnect { get; set; } = false;


        private List<string> GetAsyncList { get; set; } //Лист размеров

        //Вызов листа размеров
        public async Task<List<string>> UpdateEmployeeList()
        {
            //C# anonymous AsyncTask

            return await Task.Factory.StartNew(() =>
            {
                return GetAsyncList;
            });
        }
        //
        //Последний замер
        //
        public string lastDim { get; private set; }
        public async Task<string> AsyncLastDimmenetion()
        {
            //C# anonymous AsyncTask

            return await Task.Factory.StartNew(() =>
            {
                return lastDim;
            });
        }
        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(PropertyChangedEventArgs e)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
                handler(this, e);
        }

        protected void OnPropertyChanged(string propertyName)
        {
            OnPropertyChanged(new PropertyChangedEventArgs(propertyName));
        }

        public string LastDim
        {
            get { return lastDim; }
            set
            {
                if (value != lastDim)
                {
                    lastDim = value;
                    OnPropertyChanged("ImageFullPath");
                }
            }
        }




        public async Task<bool> UpdateConnectStat()
        {
            //C# anonymous AsyncTask

            return await Task.Factory.StartNew(() =>
            {
                return GetAsyncConnect;
            });
        }




        //Получить последний размер


        //Список всех замеров




        public void Start(string NameDevice, string Services, string Characteristic)
        {
            _services = Services;
            _characteristic = Characteristic;
            _namedevice = NameDevice;

            GetAsyncList = new List<string>();

            _watcher = new BluetoothLEAdvertisementWatcher()
            {
                ScanningMode = BluetoothLEScanningMode.Active
            };
            _watcher.Received += Watcher_Received;
            _watcher.Stopped += Watcher_Stopped;
            _watcher.Start();
        }

        public void Stop() 
        {
            _watcher.Stop();
        }
        private bool isFindDevice { get; set; } = false;
        private async void Watcher_Received(BluetoothLEAdvertisementWatcher sender, BluetoothLEAdvertisementReceivedEventArgs args)
        {
            if (isFindDevice)
                return;
            if (args.Advertisement.LocalName.Contains(_namedevice))
            {
                isFindDevice = true;
                GetAsyncConnect = true;
                BluetoothLEDevice bluetoothLeDevice = await BluetoothLEDevice.FromBluetoothAddressAsync(args.BluetoothAddress);
                GattDeviceServicesResult result = await bluetoothLeDevice.GetGattServicesAsync();//Получаем сервисы

                if (result.Status == GattCommunicationStatus.Success)
                {
                    var services = result.Services;
                    foreach (var service in services)
                    {

                        if (!service.Uuid.ToString().StartsWith(_services))
                        {
                            continue;
                        }

                        GattCharacteristicsResult characteristicsResult = await service.GetCharacteristicsAsync();

                        if (characteristicsResult.Status == GattCommunicationStatus.Success)
                        {
                            var characteristics = characteristicsResult.Characteristics;
                            foreach (var characteristic in characteristics)
                            {//Подписываемся на изменение характеристики


                                if (!characteristic.Uuid.ToString().StartsWith(_characteristic))
                                {
                                    continue;
                                }

                                GattCharacteristicProperties properties = characteristic.CharacteristicProperties;

                                if (properties.HasFlag(GattCharacteristicProperties.Notify))
                                {
                                    var notifyResult = await characteristic.WriteClientCharacteristicConfigurationDescriptorAsync(
                                          GattClientCharacteristicConfigurationDescriptorValue.Notify);
                                    if (notifyResult == GattCommunicationStatus.Success)
                                    {
                                        characteristic.ValueChanged += Charac_ValueChangedAsync;

                                        return;
                                    }
                                }

                                /* До лучших времен
if (properties.HasFlag(GattCharacteristicProperties.Read))
{
    GattReadResult gattResult = await characteristic.ReadValueAsync();
    if (gattResult.Status == GattCommunicationStatus.Success)
    {
        var reader = DataReader.FromBuffer(gattResult.Value);
        byte[] input = new byte[reader.UnconsumedBufferLength];
        reader.ReadBytes(input);

        _list.Add("value " + Encoding.UTF8.GetString(input, 0, input.Length));
        //Читаем input
    }
}

if (properties.HasFlag(GattCharacteristicProperties.Write))
{
    GattReadResult gattResult = await characteristic.WriteValueAsync(;
    if (gattResult.Status == GattCommunicationStatus.Success)
    {
        var reader = DataReader.FromBuffer(gattResult.Value);
        byte[] input = new byte[reader.UnconsumedBufferLength];
        reader.ReadBytes(input);
        //Читаем input
    }
}*/
                            }
                        }
                        else continue;

                    }
                }
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
                GetAsyncList.Add(Regex.Replace(dataFromReadResult, "[A-Za-z \n\r\0]", "").Replace('.', ','));
                LastDim = Regex.Replace(dataFromReadResult, "[A-Za-z \n\r\0]", "").Replace('.', ',');

                ///"DATA FROM READ: " + dataFromReadResult);
            }
            catch (ArgumentException)
            {
                //Debug.Write("Unknown format");
            }

        }


        private void Watcher_Stopped(BluetoothLEAdvertisementWatcher sender, BluetoothLEAdvertisementWatcherStoppedEventArgs args)
        {
            GetAsyncConnect = false;
        }
    }
}
