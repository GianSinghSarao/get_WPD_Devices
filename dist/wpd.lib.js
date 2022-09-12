//depends on JSON2.js in older scripthosts
;function get_WPD_Devices () {
  var objWMIService = new ActiveXObject("WbemScripting.SWbemLocator").ConnectServer(".", "root\\cimv2");
  var PNPDevices = objWMIService.InstancesOf("win32_PNPEntity WHERE PNPclass LIKE 'WPD'");
  var Info = new Array(PNPDevices.Count);
  for (
    var PNPDevicesList = new Enumerator(PNPDevices), 
    PNPDevice = PNPDevicesList.item(), 
    i = 0;
    !PNPDevicesList.atEnd();
    PNPDevicesList.moveNext(), 
    PNPDevice = PNPDevicesList.item(),
    i = i + 1
  ) {
    Info[i] = {};
    for (
      var props = new Enumerator(PNPDevice.Properties_),
      prop = props.item();
      !props.atEnd();
      props.moveNext(),
      prop = props.item()
    ) {
      if (prop.IsArray) {
        try {
          Info[i][prop.Name] = PNPDevice[prop.Name].toArray();
        } catch (e) {
          Info[i][prop.Name] = null;
        }
      } else {
        Info[i][prop.Name] = PNPDevice[prop.Name];
      }
    }
  }
  return Info;
};