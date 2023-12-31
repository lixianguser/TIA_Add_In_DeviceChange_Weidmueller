using System;
using Siemens.Engineering;
using Siemens.Engineering.HW;
using Siemens.Engineering.Online;
using Siemens.Engineering.AddIn.Menu;
using System.Windows;
using Siemens.Engineering.HW.Features;

namespace TIA_Add_In_DeviceChange_Weidmueller
{
    public class AddIn : ContextMenuAddIn
    {
        private static TiaPortal _tiaPortal;

        private static Device oldDevice = null;
        private static Device newDevice = null;
        private static string oldName = null;//(ReadWrite, System.String, GSD device_42)
        private static string oldPnName = null; //(ReadWrite, System.String, CC1-A5)
        private static string oldGsdName = null; //(Read, System.String, GSDML-V2.35-WI-UR20_BASIC-20210606.XML)
        private static NetworkInterface oldInterface = null;
        private static Node oldNode = null;
        private static IoSystem oldIoSystem = null;
        private static IoConnector oldIoConnector = null;
        private static Subnet oldSubnet = null;
        private static string oldIP = null;
        private static NetworkInterface newInterface = null;
        private static Node newNode = null;
        private static IoConnector newIoConnector = null;
        private static Project project;

        public AddIn(TiaPortal tiaPortal) : base("UR20_BASIC")
        {
            _tiaPortal = tiaPortal;
        }

        protected override void BuildContextMenuItems(ContextMenuAddInRoot addInRootSubmenu)
        {
            addInRootSubmenu.Items.AddActionItem<Device, DeviceItem>("Change device / version", Change_OnClick, OnClickStatus);
            //addInRootSubmenu.Items.AddActionItem<IEngineeringObject>("Change device / version",
            //    menuSelectionProvider => { }, TextInfoStatus);
        }

        private static void Change_OnClick(MenuSelectionProvider<Device, DeviceItem> menuSelectionProvider)
        {
#if DEBUG
            System.Diagnostics.Debugger.Launch();
#endif
            project = _tiaPortal.Projects[0];
            using (ExclusiveAccess exclusiveAccess = _tiaPortal.ExclusiveAccess("Change device / version..."))
            {
                using (Transaction transaction =
                       exclusiveAccess.Transaction(project, "Change device / version"))
                {
                    if (!IsOffline())
                    {
                        MessageBox.Show("PLC device not offline", "Offline Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    foreach (Device device in menuSelectionProvider.GetSelection())
                    {
                        oldDevice = device;

                        GetOldDeviceInfos();//获取旧设备信息

                        oldIoConnector.DisconnectFromIoSystem();//断开旧设备的IO系统

                        newDevice = CreateNewDevice(oldName, oldPnName);//创建一个UR20 BASIC HW2的设备

                        GetNewDeviceInfos();//获取新设备信息

                        if (oldGsdName == "GSDML-V2.35-WI-UR20_BASIC-20220715.XML")
                        {
                            //连接子网和IO系统
                            if (oldSubnet != null)
                            {
                                newNode.ConnectToSubnet(oldSubnet);
                            }
                            if (oldIoSystem != null)
                            {
                                newIoConnector.ConnectToIoSystem(oldIoSystem);
                            }
                            MoveDeviceItem();
                        }
                        else
                        {
                            //连接子网和IO系统
                            if (oldSubnet != null)
                            {
                                newNode.ConnectToSubnet(oldSubnet);
                            }
                            if (oldIoSystem != null)
                            {
                                newIoConnector.ConnectToIoSystem(oldIoSystem);
                            }
                            //创建设备项
                            CreateDeviceItem();
                        }

                        //TODO 删除旧设备
                        oldDevice.Delete();

                        //TODO 修改新设备名称
                        newDevice.Name = newDevice.Name.Replace("_COPY", string.Empty);
                        newDevice.DeviceItems[1].Name = newDevice.DeviceItems[1].Name.Replace("_COPY", string.Empty);

                        //修改IP地址
                        newNode.SetAttribute("Address", oldIP);
                    }

                    if (transaction.CanCommit)
                    {
                        transaction.CommitOnDispose();
                    }
                }
            }
        }

        /// <summary>
        /// 子网PLC是否为离线模式
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        private static bool IsOffline(/*IEngineeringServiceProvider item*/)
        {
            bool ret = false;

            foreach (Device device in project.Devices)
            {
                DeviceItem deviceItem = device.DeviceItems[1];
                if (deviceItem.GetAttribute("Classification") is DeviceItemClassifications.CPU)
                {
                    OnlineProvider onlineProvider = deviceItem.GetService<OnlineProvider>();
                    ret = (onlineProvider.State == OnlineState.Offline);
                }
            }

            return ret;
        }

        /// <summary>
        /// 判断子菜单按钮是否显示
        /// </summary>
        /// <param name="menuSelectionProvider"></param>
        /// <returns></returns>
        private static MenuStatus OnClickStatus(MenuSelectionProvider menuSelectionProvider)
        {
            foreach (Device device in menuSelectionProvider.GetSelection())
            {
                if (IsBasic(device))
                {
                    return MenuStatus.Enabled;
                }
                else
                {
                    return MenuStatus.Hidden;
                }
            }

            return MenuStatus.Disabled;
        }

        /// <summary>
        /// 创建一个新的HW2设备
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private static Device CreateNewDevice(string name, string pnName)
        {
            try
            {
                name += "_COPY";
                pnName += "_COPY";
                string gsdId = "GSD:GSDML-V2.35-WI-UR20_BASIC-20220715.XML/D";
                string pnId = "GSD:GSDML-V2.35-WI-UR20_BASIC-20220715.XML/DAP/DAP 2";
                string rackId = "GSD:GSDML-V2.35-WI-UR20_BASIC-20220715.XML/R/DAP 2";
                Device device = project.Devices.Create(gsdId, name);
                //插入插槽
                if (device.CanPlugNew(rackId, "Rack", 0))
                    device.PlugNew(rackId, "Rack", 0);

                foreach (var deviceItem in device.DeviceItems)
                {
                    //插入PN模块
                    if (deviceItem.CanPlugNew(pnId, pnName, 0))
                        deviceItem.PlugNew(pnId, pnName, 0);
                }

                return device;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        /// <summary>
        /// 判断是否为UR20模块
        /// </summary>
        /// <param name="device"></param>
        /// <returns></returns>
        private static bool IsBasic(Device device)
        {
            bool ret = false;

            try
            {
                string gsdName = (string)device.GetAttribute("GsdName");

                if (gsdName.Contains("UR20_BASIC"))
                {
                    ret = true;
                }
            }
            catch
            {
                return false;
            }

            return ret;
        }

        /// <summary>
        /// 从旧设备移动设备项到新设备
        /// </summary>
        /// <param name="oldDevice"></param>
        /// <param name="newDevice"></param>
        private static void MoveDeviceItem()
        {
            try
            {
                foreach (var oldDeviceItem in oldDevice.DeviceItems)
                {
                    int solt = (int)oldDeviceItem.GetAttribute("PositionNumber");
                    //查询是否可以在这个插槽0位置上移动模块
                    if (newDevice.DeviceItems[0].CanPlugMove(oldDeviceItem, solt))
                    {
                        newDevice.DeviceItems[0].PlugMove(oldDeviceItem, solt);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }

        /// <summary>
        /// 在新设备中创建新的设备项
        /// </summary>
        /// <param name="oldDevice"></param>
        /// <param name="newDevice"></param>
        private static void CreateDeviceItem()
        {
            try
            {
                foreach (var oldDeviceItem in oldDevice.DeviceItems)
                {
                    string name = oldDeviceItem.Name;
                    string gsdId = oldDeviceItem.GetAttribute("GsdId").ToString();
                    string gsdType = oldDeviceItem.GetAttribute("GsdType").ToString();
                    int solt = (int)oldDeviceItem.GetAttribute("PositionNumber");
                    const string newGsdId = "GSDML-V2.35-WI-UR20_BASIC-20220715.XML";
                    //int startAddress = oldDeviceItem.DeviceItems[0].Addresses[0].StartAddress;
                    //(Read, System.String, GSD:GSDML-V2.35-WI-UR20_BASIC-20210606.XML/M/ID_Mod_UR20_BASIC_16DI_P)
                    //(Read, System.String, GSD:GSDML-V2.35-WI-UR20_BASIC-20220715.XML/M/ID_Mod_UR20_BASIC_16DI_P)
                    string newDeviceItem = string.Concat("GSD:", newGsdId, "/", gsdType, "/", gsdId);
                    //查询是否可以在这个插槽0位置上移动模块
                    if (newDevice.DeviceItems[0].CanPlugNew(newDeviceItem, name, solt))
                    {
                        DeviceItem deviceItem = newDevice.DeviceItems[0].PlugNew(newDeviceItem, name, solt);
                        int oldStartAddress = oldDeviceItem.DeviceItems[0].Addresses[0].StartAddress;
                        deviceItem.DeviceItems[0].Addresses[0].SetAttribute("StartAddress", oldStartAddress);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 获取旧设备信息
        /// </summary>
        private static void GetOldDeviceInfos()
        {
            try
            {
                oldName = oldDevice.Name;
                oldPnName = oldDevice.DeviceItems[1].Name;
                oldGsdName = (string)oldDevice.GetAttribute("GsdName");

                //获取旧设备的IP地址和IO系统
                oldInterface = oldDevice.DeviceItems[1].DeviceItems[0].GetService<NetworkInterface>();
                oldNode = oldInterface.Nodes[0];
                oldIoConnector = oldInterface.IoConnectors[0];

                oldIP = oldNode.GetAttribute("Address").ToString();//192.168.0.1
                oldSubnet = oldNode.ConnectedSubnet;
                oldIoSystem = (IoSystem)oldIoConnector.GetAttribute("ConnectedToIoSystem");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 获取新设备信息
        /// </summary>
        private static void GetNewDeviceInfos()
        {
            try
            {
                //获取新设备的IP地址和IO系统
                newInterface = newDevice.DeviceItems[1].DeviceItems[0].GetService<NetworkInterface>();
                newNode = newInterface.Nodes[0];
                newIoConnector = newInterface.IoConnectors[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Exception", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}