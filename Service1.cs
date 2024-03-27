using System;
using System.ComponentModel;
using System.IO;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Threading.Tasks;
using MQTTnet;
using MQTTnet.Server;
using OfficeOpenXml;
using System.ServiceProcess;

namespace MQTT_Broker_Windows_Service
{
    public partial class Service1 : ServiceBase
    {
        protected override void OnStart(string[] args)
        {
            Task.Run(async () =>
            {
                // Get the local IP address
                string ipAddress = GetLocalIPAddress();
                Console.WriteLine($"Host IP Address: {ipAddress}");

                // Specify the port number
                int port = 1883;
                Console.WriteLine($"Port: {port}");

                var optionsBuilder = new MqttServerOptionsBuilder()
                    .WithDefaultEndpointBoundIPAddress(IPAddress.Parse(ipAddress)) // Set the local IP address for the server
                    .WithDefaultEndpointPort(port)
                    .WithConnectionValidator(c =>
                    {
                        // Accept all connections
                        c.ReasonCode = MQTTnet.Protocol.MqttConnectReasonCode.Success;

                        // Extract and print the client's IPv4 address
                        var clientIpAddress = "Unknown";
                        if (!string.IsNullOrEmpty(c.Endpoint))
                        {
                            var parts = c.Endpoint.Split(':');
                            if (parts.Length == 2 && IPAddress.TryParse(parts[0], out var clientIp) && int.TryParse(parts[1], out var clientPort))
                            {
                                var endPoint = new IPEndPoint(clientIp, clientPort);
                                clientIpAddress = endPoint.Address.ToString();
                            }
                        }
                        Console.WriteLine($"Client connected from IPv4 address: {clientIpAddress}");
                        WriteToFile($"Client connected from IPv4 address: {clientIpAddress}");
                    })
                    .WithApplicationMessageInterceptor(context =>
                    {
                        Console.WriteLine($"Client {context.ClientId} published message:");
                        Console.WriteLine($"Topic: {context.ApplicationMessage.Topic}");
                        Console.WriteLine($"Payload: {context.ApplicationMessage.ConvertPayloadToString()}");
                        SaveToExcel(context.ApplicationMessage.Topic, context.ApplicationMessage.ConvertPayloadToString());
                        WriteToFile($"Client {context.ClientId} published message: Topic: {context.ApplicationMessage.Topic} Payload: {context.ApplicationMessage.ConvertPayloadToString()}");
                    });

                var mqttServer = new MqttFactory().CreateMqttServer();

                mqttServer.ClientConnectedHandler = new MqttServerClientConnectedHandlerDelegate(e =>
                {
                    Console.WriteLine($"Client connected: {e.ClientId}");
                    WriteToFile($"Client connected: {e.ClientId}");
                });

                mqttServer.ClientDisconnectedHandler = new MqttServerClientDisconnectedHandlerDelegate(e =>
                {
                    Console.WriteLine($"Client disconnected: {e.ClientId}");
                    WriteToFile($"Client disconnected: {e.ClientId}");
                    WriteToFile("-------------------------------------------------------");
                });

                mqttServer.ClientSubscribedTopicHandler = new MqttServerClientSubscribedHandlerDelegate(e =>
                {
                    Console.WriteLine($"Client {e.ClientId} subscribed to topic: {e.TopicFilter}");
                    WriteToFile($"Client {e.ClientId} subscribed to topic: {e.TopicFilter}");
                });

                mqttServer.ClientUnsubscribedTopicHandler = new MqttServerClientUnsubscribedTopicHandlerDelegate(e =>
                {
                    Console.WriteLine($"Client {e.ClientId} unsubscribed from topic: {e.TopicFilter}");
                    WriteToFile($"Client {e.ClientId} unsubscribed from topic: {e.TopicFilter}");
                });

                try
                {
                    await mqttServer.StartAsync(optionsBuilder.Build());
                    Console.WriteLine("MQTT Broker started. Press Ctrl + C to exit...");
                    // Wait indefinitely without blocking the thread
                    await Task.Delay(-1);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred: {ex.Message}");
                    WriteToFile($"An error occurred: {ex.Message}");
                }
            });
        }


        protected override void OnStop()
        {
            // Add any cleanup code if needed
        }

        public static void SaveToExcel(string topic, string payload)
        {
            String path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filePath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\mqtt_messagges.csv";
            string[] fields = { "BMCODE", "Temperature", "Pressure", "Volume", "Level", "Generator", "Grid", "Aggregate", "Compressor1", "Compressor2", "CIP", "VoltageU", "VoltageV", "VoltageW", "CurrentU", "CurrentV", "CurrentW", "Frequency", "PwrF", "TPwr", "Time", "Date", "Topic" };
            string[] parts = payload.Split(',');

            // Combine data with topic
            string[] data = {
                parts.Length > 0 ? parts[0] : "",
                parts.Length > 1 ? parts[1] : "",
                parts.Length > 2 ? parts[2] : "",
                parts.Length > 3 ? parts[3] : "",
                parts.Length > 4 ? parts[4] : "",
                parts.Length > 5 ? parts[5] : "",
                parts.Length > 6 ? parts[6] : "",
                parts.Length > 7 ? parts[7] : "",
                parts.Length > 8 ? parts[8] : "",
                parts.Length > 9 ? parts[9] : "",
                parts.Length > 10 ? parts[10] : "",
                parts.Length > 11 ? parts[11] : "",
                parts.Length > 12 ? parts[12] : "",
                parts.Length > 13 ? parts[13] : "",
                parts.Length > 14 ? parts[14] : "",
                parts.Length > 15 ? parts[15] : "",
                parts.Length > 16 ? parts[16] : "",
                parts.Length > 17 ? parts[17] : "",
                parts.Length > 18 ? parts[18] : "",
                parts.Length > 19 ? parts[19] : "",
                parts.Length > 20 ? parts[20] : "",
                parts.Length > 21 ? parts[21] : "",
                topic
            };

            // Combine data into CSV format
            string csvRow = string.Join(",", data);
            if (!File.Exists(filePath))
            {
                // If the file doesn't exist, create it and write the header
                string header = string.Join(",", fields);
                File.WriteAllText(filePath, header + Environment.NewLine);
            }
            // Append row to CSV file
            File.AppendAllText(filePath, csvRow + Environment.NewLine);
        }

        private void WriteToFile(String Message)
        {
            String path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filePath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\mqtt_logs.txt";
            if (!File.Exists(filePath))
            {
                using (StreamWriter sw = File.CreateText(filePath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filePath))
                {
                    sw.WriteLine(Message);
                }
            }
        }

        public static string GetLocalIPAddress()
        {
            string localIP = "";
            foreach (var netInterface in NetworkInterface.GetAllNetworkInterfaces())
            {
                if (netInterface.NetworkInterfaceType == NetworkInterfaceType.Wireless80211 &&
                    netInterface.OperationalStatus == OperationalStatus.Up)
                {
                    foreach (var addrInfo in netInterface.GetIPProperties().UnicastAddresses)
                    {
                        if (addrInfo.Address.AddressFamily == AddressFamily.InterNetwork)
                        {
                            localIP = addrInfo.Address.ToString();
                        }
                    }
                }
            }
            return localIP;
        }
    }
}
