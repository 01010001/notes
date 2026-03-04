import ssl
import logging
from pyVim.connect import SmartConnect, Disconnect
from pyVmomi import vim
import pandas as pd
from datetime import datetime
import ipaddress
 
# Log yapılandırması
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
 
# Birden fazla vCenter bilgisi
vcenters = [
        {"ip": "10.1.28.200", "user":   "adm1711", "password": "Sinan98329832*-"},
        {"ip": "10.1.28.100", "user":   "adm1711", "password": "Sinan98329832*-"},
        {"ip": "10.150.37.100", "user":  "adm1711", "password": "Sinan98329832*-"},
        {"ip": "10.150.37.200", "user": "adm1711", "password": "Sinan98329832*-"}
]
 
# SSL sertifikalarını kontrol etmemek için
context = ssl._create_unverified_context()
 
def connect_vcenter(vc):
    """vCenter'a bağlanma fonksiyonu."""
    try:
        si = SmartConnect(host=vc["ip"], user=vc["user"], pwd=vc["password"], sslContext=context)
        logging.info(f"{vc['ip']} adresine bağlantı başarılı.")
        return si
    except Exception as e:
        logging.error(f"{vc['ip']} adresine bağlanırken hata oluştu: {e}")
        return None
 
def get_vm_info(vm, datacenter_name, folder_name, allowed_guests=[]):
    """Bir sanal makinenin bilgilerini döner."""
    try:
        config = vm.config
        datastores = [ds.info.name for ds in vm.datastore] if vm.datastore else None
        disk_size_gb = sum(
            device.capacityInKB / 1024**2 for device in config.hardware.device if isinstance(device, vim.vm.device.VirtualDisk)
        )
 
        guest_full_name = vm.guest.guestFullName if vm.guest else None
 
        # allowed_guests boşsa filtre devre dışı kalır
        if not allowed_guests or (guest_full_name and any(guest in guest_full_name for guest in allowed_guests)):
            return {
                "Datacenter": datacenter_name,
                "Folder": folder_name,
                "Name": vm.name,
                "Hostname": vm.guest.hostName if vm.guest else None,
                "IPv4 Address": vm.guest.ipAddress,
                "Power Status": vm.runtime.powerState,
                "Guest Full Name": guest_full_name,
                "Datastore": ', '.join(datastores) if datastores else None,
                "vCPU": config.hardware.numCPU if config else None,
                "Memory (MB)": config.hardware.memoryMB if config else None,
                "Disk Size (GB)": disk_size_gb,
                "Notes": vm.config.annotation,
            }
    except Exception as e:
        logging.warning(f"{vm.name} VM bilgileri alınırken hata oluştu: {e}")
 
    return None
 
def get_all_vms(datacenter):
    """Datacenter içindeki tüm VM'leri rekürsif olarak bulur."""
    def get_vms_in_entity(entity, datacenter_name, folder_name=None):
        vms = []
        for item in entity.childEntity:
            if isinstance(item, vim.VirtualMachine):
                vms.append((item, datacenter_name, folder_name))
            elif isinstance(item, vim.Folder):
                vms.extend(get_vms_in_entity(item, datacenter_name, item.name))
        return vms
 
    return get_vms_in_entity(datacenter.vmFolder, datacenter.name)
 
def fetch_data_from_vcenters(allowed_guests=[]):
    """Her bir vCenter'dan VM verilerini çeker ve liste olarak döner."""
    all_data = []
    for index, vc in enumerate(vcenters):
        si = connect_vcenter(vc)
        if si:
            try:
                for datacenter in si.content.rootFolder.childEntity:
                    if isinstance(datacenter, vim.Datacenter):
                        for vm, datacenter_name, folder_name in get_all_vms(datacenter):
                            vm_info = get_vm_info(vm, datacenter_name, folder_name, allowed_guests)
                            if vm_info:
                                vm_info["vCenter"] = vc["ip"]
                                all_data.append(vm_info)  # Tüm verileri ekle
            finally:
                Disconnect(si)
                logging.info(f"{vc['ip']} adresine bağlantı kapatıldı.")
    return all_data
 
def load_subnets(filenames):
    """Bir veya daha fazla dosyadan subnet bilgilerini okur ve döner."""
    subnets = []
    for filename in filenames:
        try:
            with open(filename, 'r') as file:
                # CIDR notasyonunu kullanarak subnetleri yükle
                subnets.extend([ipaddress.ip_network(line.strip(), strict=False) for line in file if line.strip()])
            logging.info(f"{filename} dosyasından subnet bilgileri yüklendi: {subnets}")
        except Exception as e:
            logging.error(f"{filename} dosyası okunurken hata oluştu: {e}")
    return subnets
 
def save_to_excel(data, subnet_list1, subnet_list2, subnet_list3, subnet_list4):
    """Verileri Excel dosyasına kaydeder."""
    file_name = f"vcenter_virtual_machines_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
 
    # İlk alt ağ dosyası için gruplama
    subnet_data1 = []
    for vm in data:
        ipv4_address = vm.get("IPv4 Address")
        if ipv4_address:
            vm_ip = ipaddress.ip_address(ipv4_address)
            if any(vm_ip in subnet for subnet in subnet_list1):
                subnet_data1.append(vm)  # Sadece belirtilen subnetlerdeki VM'leri ekle
 
    # İkinci alt ağ dosyası için gruplama
    subnet_data2 = []
    for vm in data:
        ipv4_address = vm.get("IPv4 Address")
        if ipv4_address:
            vm_ip = ipaddress.ip_address(ipv4_address)
            if any(vm_ip in subnet for subnet in subnet_list2):
                subnet_data2.append(vm)  # Sadece belirtilen subnetlerdeki VM'leri ekle
 
  # İlk alt ağ dosyası için gruplama
    subnet_data3 = []
    for vm in data:
        ipv4_address = vm.get("IPv4 Address")
        if ipv4_address:
            vm_ip = ipaddress.ip_address(ipv4_address)
            if any(vm_ip in subnet for subnet in subnet_list3):
                subnet_data3.append(vm)  # Sadece belirtilen subnetlerdeki VM'leri ekle
 
  # İlk alt ağ dosyası için gruplama
    subnet_data4 = []
    for vm in data:
        ipv4_address = vm.get("IPv4 Address")
        if ipv4_address:
            vm_ip = ipaddress.ip_address(ipv4_address)
            if any(vm_ip in subnet for subnet in subnet_list4):
                subnet_data4.append(vm)  # Sadece belirtilen subnetlerdeki VM'leri ekle
 
    # Excel dosyasına yazma
    with pd.ExcelWriter(file_name) as writer:
        # İlk sayfaya tüm sunucuları ekle
        df_all_vms = pd.DataFrame(data)
        df_all_vms.to_excel(writer, sheet_name='All_Virtual_Machines', index=False)
 
        # İlk subnet dosyasındaki sunucuları ekle
        df_subnets1 = pd.DataFrame(subnet_data1)
        df_subnets1.to_excel(writer, sheet_name='Istanbul', index=False) #Sheet isimlerini değiştirmek için
 
        # İkinci subnet dosyasındaki sunucuları ekle
        df_subnets2 = pd.DataFrame(subnet_data2)
        df_subnets2.to_excel(writer, sheet_name='Istanbul-DMZ', index=False) #İkinci Sheet ismini değiştirmek için
 
        # İkinci subnet dosyasındaki sunucuları ekle
        df_subnets3 = pd.DataFrame(subnet_data3)
        df_subnets3.to_excel(writer, sheet_name='Ankara', index=False) #İkinci Sheet ismini değiştirmek için
 
        # İkinci subnet dosyasındaki sunucuları ekle
        df_subnets4 = pd.DataFrame(subnet_data4)
        df_subnets4.to_excel(writer, sheet_name='Ankara-DMZ', index=False) #İkinci Sheet ismini değiştirmek için
 
    logging.info(f"Excel dosyası başarıyla oluşturuldu: {file_name}")
 
if __name__ == "__main__":
    allowed_guests = ["Linux","linux"]  # Filtrelemek istediğin guest türlerini buraya ekle
    data = fetch_data_from_vcenters(allowed_guests)
    subnet_list1 = load_subnets(['istanbul.txt'])  # İlk subnet dosyasını yükle
    subnet_list2 = load_subnets(['ist-dmz.txt'])  # İkinci subnet dosyasını yükle
    subnet_list3 = load_subnets(['ankara.txt'])  # ücüncü subnet dosyasını yükle
    subnet_list4 = load_subnets(['ank-dmz.txt'])  # dördüncü subnet dosyasını yükle
    if data:
        save_to_excel(data, subnet_list1, subnet_list2, subnet_list3, subnet_list4)