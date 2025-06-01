using ExcelDataReader;
using PhanMemQuanLyTaiSan.model;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Windows;
using System.Windows.Input;

namespace PhanMemQuanLyTaiSan.viewmodel
{
    public class MainViewModel : INotifyPropertyChanged
    {
        public ObservableCollection<CoSo> Facility { get; set; } = new ObservableCollection<CoSo>();
        public ObservableCollection<KhongGian> Space { get; set; } = new ObservableCollection<KhongGian>();
        public ObservableCollection<LienHe> Contact { get; set; } = new ObservableCollection<LienHe>();
        public ObservableCollection<San> Floor { get; set; } = new ObservableCollection<San>();
        public ObservableCollection<KhuVuc> Zone { get; set; } = new ObservableCollection<KhuVuc>();
        public ObservableCollection<Loai> Type { get; set; } = new ObservableCollection<Loai>();
        public ObservableCollection<ThanhPhan> Component { get; set; } = new ObservableCollection<ThanhPhan>();
        public ObservableCollection<HeThong> Systems { get; set; } = new ObservableCollection<HeThong>();
        public ObservableCollection<ThuocTinh> Attributes { get; set; } = new ObservableCollection<ThuocTinh>();
        public ObservableCollection<DieuPhoi> Coordinates { get; set; } = new ObservableCollection<DieuPhoi>();

        public ICommand LoadExcelCommand { get; }
        public ICommand DeleteContactCommand { get; }
        public ICommand DeleteFacilityCommand { get; }
        public ICommand DeleteSpaceCommand { get; }
        public ICommand DeleteFloorCommand { get; }
        public ICommand DeleteZoneCommand { get; }
        public ICommand DeleteTypeCommand { get; }
        public ICommand DeleteComponentCommand { get; }
        public ICommand DeleteSystemsCommand { get; }
        public ICommand DeleteAttributesCommand { get; }
        public ICommand DeleteCoordinatesCommand { get; }
        public ICommand AddContactCommand { get; }
        public ICommand AddFacilityCommand { get; }
        public ICommand AddSpaceCommand { get; }
        public ICommand AddFloorCommand { get; }
        public ICommand AddZoneCommand { get; }
        public ICommand AddTypeCommand { get; }
        public ICommand AddComponentCommand { get; }
        public ICommand AddSystemsCommand { get; }
        public ICommand AddAttributesCommand { get; }
        public ICommand AddCoordinatesCommand { get; }


        public MainViewModel()
        {
            MessageBox.Show("Đăng nhập thành công, mời bạn chọn file excel!");
            var filePath = SelectExcelFile();
            if (!string.IsNullOrEmpty(filePath))
            {
                LoadFromExcel(filePath);
            }
            DeleteContactCommand = new RelayCommand(_ => DeleteSelectedContact(), _ => SelectedContact != null);
            DeleteFacilityCommand = new RelayCommand(_ => DeleteSelectedFacility(), _ => SelectedFacility != null);
            DeleteSpaceCommand = new RelayCommand(_ => DeleteSelectedSpace(), _ => SelectedSpace != null);
            DeleteFloorCommand = new RelayCommand(_ => DeleteSelectedFloor(), _ => SelectedFloor != null);
            DeleteZoneCommand = new RelayCommand(_ => DeleteSelectedZone(), _ => SelectedZone != null);
            DeleteTypeCommand = new RelayCommand(_ => DeleteSelectedType(), _ => SelectedType != null);
            DeleteComponentCommand = new RelayCommand(_ => DeleteSelectedComponent(), _ => SelectedComponent != null);
            DeleteSystemsCommand = new RelayCommand(_ => DeleteSelectedSystems(), _ => SelectedSystems != null);
            DeleteAttributesCommand = new RelayCommand(_ => DeleteSelectedAttributes(), _ => SelectedAttributes != null);
            DeleteCoordinatesCommand = new RelayCommand(_ => DeleteSelectedCoordinates(), _ => SelectedCoordinates != null);
            AddContactCommand = new RelayCommand(_ => AddNewContact());
            AddFacilityCommand = new RelayCommand(_ => AddNewFacility());
            AddSpaceCommand = new RelayCommand(_ => AddNewSpace());
            AddFloorCommand = new RelayCommand(_ => AddNewFloor());
            AddZoneCommand = new RelayCommand(_ => AddNewZone());
            AddTypeCommand = new RelayCommand(_ => AddNewType());
            AddComponentCommand = new RelayCommand(_ => AddNewComponent());
            AddSystemsCommand = new RelayCommand(_ => AddNewSystems());
            AddAttributesCommand = new RelayCommand(_ => AddNewAttributes());
            AddCoordinatesCommand = new RelayCommand(_ => AddNewCoordinates());

        }

        private void DeleteSelectedContact()
        {
            if (SelectedContact != null)
            {
                Contact.Remove(SelectedContact);
            }
        }
        private LienHe _selectedContact;
        private void DeleteSelectedFacility()
        {
            if (SelectedFacility != null)
            {
                Facility.Remove(SelectedFacility);
            }
        }

        private void DeleteSelectedSpace()
        {
            if (SelectedSpace != null)
            {
                Space.Remove(SelectedSpace);
            }
        }

        private void DeleteSelectedFloor()
        {
            if (SelectedFloor != null)
            {
                Floor.Remove(SelectedFloor);
            }
        }

        private void DeleteSelectedZone()
        {
            if (SelectedZone != null)
            {
                Zone.Remove(SelectedZone);
            }
        }

        private void DeleteSelectedType()
        {
            if (SelectedType != null)
            {
                Type.Remove(SelectedType);
            }
        }

        private void DeleteSelectedComponent()
        {
            if (SelectedComponent != null)
            {
                Component.Remove(SelectedComponent);
            }
        }

        private void DeleteSelectedSystems()
        {
            if (SelectedSystems != null)
            {
                Systems.Remove(SelectedSystems);
            }
        }

        private void DeleteSelectedAttributes()
        {
            if (SelectedAttributes != null)
            {
                Attributes.Remove(SelectedAttributes);
            }
        }

        private void DeleteSelectedCoordinates()
        {
            if (SelectedCoordinates != null)
            {
                Coordinates.Remove(SelectedCoordinates);
            }
        }

        public LienHe SelectedContact
        {
            get => _selectedContact;
            set
            {
                _selectedContact = value;
                OnPropertyChanged(nameof(SelectedContact));
            }
        }
        private CoSo _selectedFacility;
        public CoSo SelectedFacility
        {
            get => _selectedFacility;
            set
            {
                _selectedFacility = value;
                OnPropertyChanged(nameof(SelectedFacility));
            }
        }

        private KhongGian _selectedSpace;
        public KhongGian SelectedSpace
        {
            get => _selectedSpace;
            set
            {
                _selectedSpace = value;
                OnPropertyChanged(nameof(SelectedSpace));
            }
        }

        private San _selectedFloor;
        public San SelectedFloor
        {
            get => _selectedFloor;
            set
            {
                _selectedFloor = value;
                OnPropertyChanged(nameof(SelectedFloor));
            }
        }

        private KhuVuc _selectedZone;
        public KhuVuc SelectedZone
        {
            get => _selectedZone;
            set
            {
                _selectedZone = value;
                OnPropertyChanged(nameof(SelectedZone));
            }
        }

        private Loai _selectedType;
        public Loai SelectedType
        {
            get => _selectedType;
            set
            {
                _selectedType = value;
                OnPropertyChanged(nameof(SelectedType));
            }
        }

        private ThanhPhan _selectedComponent;
        public ThanhPhan SelectedComponent
        {
            get => _selectedComponent;
            set
            {
                _selectedComponent = value;
                OnPropertyChanged(nameof(SelectedComponent));
            }
        }

        private HeThong _selectedSystems;
        public HeThong SelectedSystems
        {
            get => _selectedSystems;
            set
            {
                _selectedSystems = value;
                OnPropertyChanged(nameof(SelectedSystems));
            }
        }

        private ThuocTinh _selectedAttributes;
        public ThuocTinh SelectedAttributes
        {
            get => _selectedAttributes;
            set
            {
                _selectedAttributes = value;
                OnPropertyChanged(nameof(SelectedAttributes));
            }
        }

        private DieuPhoi _selectedCoordinates;
        public DieuPhoi SelectedCoordinates
        {
            get => _selectedCoordinates;
            set
            {
                _selectedCoordinates = value;
                OnPropertyChanged(nameof(SelectedCoordinates));
            }
        }

        private void AddNewContact()
        {
            var newItem = new LienHe
            {
                Email = "email@example.com",
                CreatedBy = "Người tạo",
                CreatedOn = DateTime.Now.ToString("dd/MM/yyyy"),
                Category = "Loại liên hệ",
                Company = "Tên công ty",
                Phone = "0123456789",
                ExtSystem = "Hệ thống ngoài",
                ExtObject = "Đối tượng liên hệ",
                ExtIdentifier = "Mã định danh",
                Department = "Phòng ban",
                OrganizationCode = "Mã tổ chức",
                GivenName = "Tên",
                FamilyName = "Họ",
                Street = "Số nhà, Đường",
                PostalBox = "",
                Town = "Thị trấn/Quận",
                StateRegion = "Tỉnh/Thành phố",
                PostalCode = "100000",
                Country = "Việt Nam"
            };

            Contact.Add(newItem);
            SelectedContact = newItem;
        }
        private void AddNewFacility()
        {
            var newItem = new CoSo
            {
                Name = "Tên cơ sở",
                CreatedBy = "Người tạo",
                CreatedOn = DateTime.Now.ToString("dd/MM/yyyy"),
                Category = "Loại",
                ProjectName = "Tên dự án",
                SiteName = "Tên địa điểm",
                LinearUnits = "m",
                AreaUnits = "m2",
                VolumeUnits = "m3",
                CurrencyUnit = "VND",
                AreaMeasurement = "Diện tích đo được",
                ExtSystem = "Hệ thống ngoài",
                ExtProjectObject = "Đối tượng dự án",
                ExtProjectIdentifier = "Mã dự án",
                ExtSiteObject = "Đối tượng địa điểm",
                ExtSiteIdentifier = "Mã địa điểm",
                ExtFacilityObject = "Đối tượng cơ sở",
                ExtFacilityIdentifier = "Mã cơ sở",
                Description = "Mô tả",
                ProjectDescription = "Mô tả dự án",
                SiteDescription = "Mô tả địa điểm",
                Phase = "Giai đoạn"
            };

            Facility.Add(newItem);
            SelectedFacility = newItem;
        }
        private void AddNewSpace()
        {
            var newItem = new KhongGian
            {
                Name = "Tên không gian",
                CreatedBy = "Người tạo",
                CreatedOn = DateTime.Now.ToString("dd/MM/yyyy"),
                Category = "Loại",
                FloorName = "Tên tầng",
                Description = "Mô tả",
                ExtSystem = "Hệ thống ngoài",
                ExtObject = "Đối tượng",
                ExtIdentifier = "Mã định danh",
                UsableHeight = "Chiều cao sử dụng",
                GrossArea = "Tổng diện tích",
                NetArea = "Diện tích sử dụng"
            };

            Space.Add(newItem);
            SelectedSpace = newItem;
        }
        private void AddNewFloor()
        {
            var newItem = new San
            {
                Name = "Tên sàn",
                CreatedBy = "Người tạo",
                CreatedOn = DateTime.Now.ToString("dd/MM/yyyy"),
                Category = "Loại",
                ExtSystem = "Hệ thống ngoài",
                ExtObject = "Đối tượng",
                ExtIdentifier = "Mã định danh",
                Description = "Mô tả",
                Elevation = "Độ cao",
                Height = "Chiều cao"
            };

            Floor.Add(newItem);
            SelectedFloor = newItem;
        }
        private void AddNewZone()
        {
            var newItem = new KhuVuc
            {
                Name = "Tên khu vực",
                CreatedBy = "Người tạo",
                CreatedOn = DateTime.Now.ToString("dd/MM/yyyy"),
                Category = "Loại",
                SpaceNames = "Tên không gian",
                ExtSystem = "Hệ thống ngoài",
                ExtObject = "Đối tượng",
                ExtIdentifier = "Mã định danh",
                Description = "Mô tả"
            };

            Zone.Add(newItem);
            SelectedZone = newItem;
        }
        private void AddNewType()
        {
            var newItem = new Loai
            {
                Name = "Tên loại",
                CreatedBy = "Người tạo",
                CreatedOn = DateTime.Now.ToString("dd/MM/yyyy"),
                Category = "Danh mục",
                Description = "Mô tả",
                AssetType = "Thiết bị",
                Manufacturer = "Nhà sản xuất",
                ModelNumber = "Model123",
                WarrantyGuarantorParts = "Bảo hành bộ phận",
                WarrantyDurationParts = "12 tháng",
                WarrantyGuarantorLabor = "Bảo hành nhân công",
                WarrantyDurationUnit = "Tháng",
                ExtSystem = "Hệ thống",
                ExtObject = "Đối tượng",
                ExtIdentifier = "Mã123",
                ReplacementCost = "1000000",
                ExpectedLife = "5 năm",
                DurationUnit = "Năm",
                WarrantyDescription = "Mô tả bảo hành",
                NominalLength = "100",
                NominalWidth = "50",
                NominalHeight = "25",
                ModelReference = "RefModel",
                Shape = "Hình dạng",
                Size = "Kích thước",
                Color = "Màu sắc",
                Grade = "Cấp",
                Material = "Vật liệu",
                Constituents = "Thành phần",
                Features = "Tính năng",
                AccessibilityPerformance = "Hiệu suất tiếp cận",
                CodePerformance = "Hiệu suất mã",
                SustainabilityPerformance = "Hiệu suất bền vững",
                Area = "200",
                Length = "300"
            };

            Type.Add(newItem);
            SelectedType = newItem;
        }
        private void AddNewComponent()
        {
            var newItem = new ThanhPhan
            {
                Name = "Tên thành phần",
                CreatedBy = "Người tạo",
                CreatedOn = DateTime.Now.ToString("dd/MM/yyyy"),
                TypeName = "Kiểu",
                Space = "Không gian",
                Description = "Mô tả",
                ExtSystem = "Hệ thống",
                ExtObject = "Đối tượng",
                ExtIdentifier = "Mã định danh",
                SerialNumber = "SN123",
                InstallationDate = DateTime.Now.ToString("dd/MM/yyyy"),
                WarrantyStartDate = DateTime.Now.ToString("dd/MM/yyyy"),
                TagNumber = "Tag001",
                AssetIdentifier = "AssetID001",
                Area = "100",
                Length = "200"
            };

            Component.Add(newItem);
            SelectedComponent = newItem;
        }
        private void AddNewSystems()
        {
            var newItem = new HeThong
            {
                Name = "Tên hệ thống",
                CreatedBy = "Người tạo",
                CreatedOn = DateTime.Now.ToString("dd/MM/yyyy"),
                Category = "Loại",
                ComponentNames = "Danh sách thành phần",
                ExtSystem = "Hệ thống ngoài",
                ExtObject = "Đối tượng",
                ExtIdentifier = "Mã định danh",
                Description = "Mô tả"
            };

            Systems.Add(newItem);
            SelectedSystems = newItem;
        }
        private void AddNewAttributes()
        {
            var newItem = new ThuocTinh
            {
                Name = "Tên thuộc tính",
                CreatedBy = "Người tạo",
                CreatedOn = DateTime.Now.ToString("dd/MM/yyyy"),
                Category = "Loại",
                SheetName = "Tên sheet",
                RowName = "Tên hàng",
                Value = "Giá trị",
                Unit = "Đơn vị",
                ExtSystem = "Hệ thống ngoài",
                ExtObject = "Đối tượng",
                ExtIdentifier = "Mã định danh",
                Description = "Mô tả",
                AllowedValues = "Giá trị cho phép"
            };

            Attributes.Add(newItem);
            SelectedAttributes = newItem;
        }
        private void AddNewCoordinates()
        {
            var newItem = new DieuPhoi
            {
                Name = "Tên điều phối",
                CreatedBy = "Người tạo",
                CreatedOn = DateTime.Now.ToString("dd/MM/yyyy"),
                Category = "Loại",
                SheetName = "Tên sheet",
                RowName = "Tên hàng",
                CoordinateXAxis = "0.0",
                CoordinateYAxis = "0.0",
                CoordinateZAxis = "0.0",
                ExtSystem = "Hệ thống ngoài",
                ExtObject = "Đối tượng",
                ExtIdentifier = "Mã định danh",
                ClockwiseRotation = "0",
                ElevationalRotation = "0",
                YawRotation = "0"
            };

            Coordinates.Add(newItem);
            SelectedCoordinates = newItem;
        }

        private string SelectExcelFile()
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Chọn file Excel",
                Filter = "Excel Files|*.xlsx;*.xls",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };

            bool? result = dialog.ShowDialog();

            return result == true ? dialog.FileName : null;
        }
        private DataTable ReadSheetToDataTable(string filePath, string sheetName)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var conf = new ExcelDataSetConfiguration
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = true }
                    };

                    var dataSet = reader.AsDataSet(conf);
                    var table = dataSet.Tables[sheetName];
                    if (table == null)
                        throw new Exception($"Không tìm thấy sheet '{sheetName}' trong file Excel.");

                    return table;
                }
            }
        }

        private void LoadFromExcel(string filePath)
        {
            try
            {
                // Facility
                var dtFacility = ReadSheetToDataTable(filePath, "Facility");
                Facility.Clear();
                string[] facilityColumns = new string[] {
                    "Name", "CreatedBy", "CreatedOn", "Category", "ProjectName", "SiteName",
                    "LinearUnits", "AreaUnits", "VolumeUnits", "CurrencyUnit", "AreaMeasurement",
                    "ExtSystem", "ExtProjectObject", "ExtProjectIdentifier", "ExtSiteObject",
                    "ExtSiteIdentifier", "ExtFacilityObject", "ExtFacilityIdentifier", "Description",
                    "ProjectDescription", "SiteDescription", "Phase"
                };
                CheckColumnsExist(dtFacility, "Facility", facilityColumns);
                foreach (DataRow row in dtFacility.Rows)
                {
                    Facility.Add(new CoSo
                    {
                        Name = row["Name"].ToString(),
                        CreatedBy = row["CreatedBy"].ToString(),
                        CreatedOn = row["CreatedOn"].ToString(),
                        Category = row["Category"].ToString(),
                        ProjectName = row["ProjectName"].ToString(),
                        SiteName = row["SiteName"].ToString(),
                        LinearUnits = row["LinearUnits"].ToString(),
                        AreaUnits = row["AreaUnits"].ToString(),
                        VolumeUnits = row["VolumeUnits"].ToString(),
                        CurrencyUnit = row["CurrencyUnit"].ToString(),
                        AreaMeasurement = row["AreaMeasurement"].ToString(),
                        ExtSystem = row["ExtSystem"].ToString(),
                        ExtProjectObject = row["ExtProjectObject"].ToString(),
                        ExtProjectIdentifier = row["ExtProjectIdentifier"].ToString(),
                        ExtSiteObject = row["ExtSiteObject"].ToString(),
                        ExtSiteIdentifier = row["ExtSiteIdentifier"].ToString(),
                        ExtFacilityObject = row["ExtFacilityObject"].ToString(),
                        ExtFacilityIdentifier = row["ExtFacilityIdentifier"].ToString(),
                        Description = row["Description"].ToString(),
                        ProjectDescription = row["ProjectDescription"].ToString(),
                        SiteDescription = row["SiteDescription"].ToString(),
                        Phase = row["Phase"].ToString(),
                    });
                }

                // Space
                var dtSpace = ReadSheetToDataTable(filePath, "Space");
                Space.Clear();
                string[] spaceColumns = new string[] {
                    "Name", "CreatedBy", "CreatedOn", "Category", "FloorName", "Description",
                    "ExtSystem", "ExtObject", "ExtIdentifier", "UsableHeight", "GrossArea", "NetArea"
                };
                CheckColumnsExist(dtSpace, "Space", spaceColumns);
                foreach (DataRow row in dtSpace.Rows)
                {
                    Space.Add(new KhongGian
                    {
                        Name = row["Name"].ToString(),
                        CreatedBy = row["CreatedBy"].ToString(),
                        CreatedOn = row["CreatedOn"].ToString(),
                        Category = row["Category"].ToString(),
                        FloorName = row["FloorName"].ToString(),
                        Description = row["Description"].ToString(),
                        ExtSystem = row["ExtSystem"].ToString(),
                        ExtObject = row["ExtObject"].ToString(),
                        ExtIdentifier = row["ExtIdentifier"].ToString(),
                        UsableHeight = row["UsableHeight"].ToString(),
                        GrossArea = row["GrossArea"].ToString(),
                        NetArea = row["NetArea"].ToString(),
                    });
                }

                // Contact
                var dtContacts = ReadSheetToDataTable(filePath, "Contact");
                Contact.Clear();
                string[] contactColumns = new string[] {
                    "Email", "CreatedBy", "CreatedOn", "Category", "Company", "Phone", "ExtSystem",
                    "ExtObject", "ExtIdentifier", "Department", "OrganizationCode", "GivenName",
                    "FamilyName", "Street", "PostalBox", "Town", "StateRegion", "PostalCode", "Country"
                };
                CheckColumnsExist(dtContacts, "Contact", contactColumns);
                foreach (DataRow row in dtContacts.Rows)
                {
                    Contact.Add(new LienHe
                    {
                        Email = row["Email"].ToString(),
                        CreatedBy = row["CreatedBy"].ToString(),
                        CreatedOn = row["CreatedOn"].ToString(),
                        Category = row["Category"].ToString(),
                        Company = row["Company"].ToString(),
                        Phone = row["Phone"].ToString(),
                        ExtSystem = row["ExtSystem"].ToString(),
                        ExtObject = row["ExtObject"].ToString(),
                        ExtIdentifier = row["ExtIdentifier"].ToString(),
                        Department = row["Department"].ToString(),
                        OrganizationCode = row["OrganizationCode"].ToString(),
                        GivenName = row["GivenName"].ToString(),
                        FamilyName = row["FamilyName"].ToString(),
                        Street = row["Street"].ToString(),
                        PostalBox = row["PostalBox"].ToString(),
                        Town = row["Town"].ToString(),
                        StateRegion = row["StateRegion"].ToString(),
                        PostalCode = row["PostalCode"].ToString(),
                        Country = row["Country"].ToString()
                    });
                }

                // Floor
                var dtFloor = ReadSheetToDataTable(filePath, "Floor");
                Floor.Clear();
                string[] floorColumns = new string[] {
                    "Name", "CreatedBy", "CreatedOn", "Category", "ExtSystem", "ExtObject",
                    "ExtIdentifier", "Description", "Elevation", "Height"
                };
                CheckColumnsExist(dtFloor, "Floor", floorColumns);
                foreach (DataRow row in dtFloor.Rows)
                {
                    Floor.Add(new San
                    {
                        Name = row["Name"].ToString(),
                        CreatedBy = row["CreatedBy"].ToString(),
                        CreatedOn = row["CreatedOn"].ToString(),
                        Category = row["Category"].ToString(),
                        ExtSystem = row["ExtSystem"].ToString(),
                        ExtObject = row["ExtObject"].ToString(),
                        ExtIdentifier = row["ExtIdentifier"].ToString(),
                        Description = row["Description"].ToString(),
                        Elevation = row["Elevation"].ToString(),
                        Height = row["Height"].ToString()
                    });
                }

                // Zone
                var dtZone = ReadSheetToDataTable(filePath, "Zone");
                Zone.Clear();
                string[] zoneColumns = new string[] {
                    "Name", "CreatedBy", "CreatedOn", "Category", "SpaceNames", "ExtSystem",
                    "ExtObject", "ExtIdentifier", "Description"
                };
                CheckColumnsExist(dtZone, "Zone", zoneColumns);
                foreach (DataRow row in dtZone.Rows)
                {
                    Zone.Add(new KhuVuc
                    {
                        Name = row["Name"].ToString(),
                        CreatedBy = row["CreatedBy"].ToString(),
                        CreatedOn = row["CreatedOn"].ToString(),
                        Category = row["Category"].ToString(),
                        SpaceNames = row["SpaceNames"].ToString(),
                        ExtSystem = row["ExtSystem"].ToString(),
                        ExtObject = row["ExtObject"].ToString(),
                        ExtIdentifier = row["ExtIdentifier"].ToString(),
                        Description = row["Description"].ToString()
                    });
                }

                // Type
                var dtType = ReadSheetToDataTable(filePath, "Type");
                Type.Clear();
                string[] typeColumns = new string[]
                {
                    "Name", "CreatedBy", "CreatedOn", "Category", "Description", "AssetType", "Manufacturer", "ModelNumber",
                    "WarrantyGuarantorParts", "WarrantyDurationParts", "WarrantyGuarantorLabor", "WarrantyDurationUnit",
                    "ExtSystem", "ExtObject", "ExtIdentifier", "ReplacementCost", "ExpectedLife", "DurationUnit",
                    "WarrantyDescription", "NominalLength", "NominalWidth", "NominalHeight", "ModelReference", "Shape",
                    "Size", "Color", "Grade", "Material", "Constituents", "Features", "AccessibilityPerformance",
                    "CodePerformance", "SustainabilityPerformance", "Area", "Length"
                };
                CheckColumnsExist(dtType, "Type", typeColumns);
                foreach (DataRow row in dtType.Rows)
                {
                    Type.Add(new Loai
                    {
                        Name = row["Name"].ToString(),
                        CreatedBy = row["CreatedBy"].ToString(),
                        CreatedOn = row["CreatedOn"].ToString(),
                        Category = row["Category"].ToString(),
                        Description = row["Description"].ToString(),
                        AssetType = row["AssetType"].ToString(),
                        Manufacturer = row["Manufacturer"].ToString(),
                        ModelNumber = row["ModelNumber"].ToString(),
                        WarrantyGuarantorParts = row["WarrantyGuarantorParts"].ToString(),
                        WarrantyDurationParts = row["WarrantyDurationParts"].ToString(),
                        WarrantyGuarantorLabor = row["WarrantyGuarantorLabor"].ToString(),
                        WarrantyDurationUnit = row["WarrantyDurationUnit"].ToString(),
                        ExtSystem = row["ExtSystem"].ToString(),
                        ExtObject = row["ExtObject"].ToString(),
                        ExtIdentifier = row["ExtIdentifier"].ToString(),
                        ReplacementCost = row["ReplacementCost"].ToString(),
                        ExpectedLife = row["ExpectedLife"].ToString(),
                        DurationUnit = row["DurationUnit"].ToString(),
                        WarrantyDescription = row["WarrantyDescription"].ToString(),
                        NominalLength = row["NominalLength"].ToString(),
                        NominalWidth = row["NominalWidth"].ToString(),
                        NominalHeight = row["NominalHeight"].ToString(),
                        ModelReference = row["ModelReference"].ToString(),
                        Shape = row["Shape"].ToString(),
                        Size = row["Size"].ToString(),
                        Color = row["Color"].ToString(),
                        Grade = row["Grade"].ToString(),
                        Material = row["Material"].ToString(),
                        Constituents = row["Constituents"].ToString(),
                        Features = row["Features"].ToString(),
                        AccessibilityPerformance = row["AccessibilityPerformance"].ToString(),
                        CodePerformance = row["CodePerformance"].ToString(),
                        SustainabilityPerformance = row["SustainabilityPerformance"].ToString(),
                        Area = row["Area"].ToString(),
                        Length = row["Length"].ToString()
                    });
                }

                // Component
                var dtComponent = ReadSheetToDataTable(filePath, "Component");
                Component.Clear();
                string[] componentColumns = new string[]
                {
                    "Name", "CreatedBy", "CreatedOn", "TypeName", "Space", "Description",
                    "ExtSystem", "ExtObject", "ExtIdentifier", "SerialNumber", "InstallationDate",
                    "WarrantyStartDate", "TagNumber", "AssetIdentifier", "Area", "Length"
                };
                CheckColumnsExist(dtComponent, "Component", componentColumns);
                foreach (DataRow row in dtComponent.Rows)
                {
                    Component.Add(new ThanhPhan
                    {
                        Name = row["Name"].ToString(),
                        CreatedBy = row["CreatedBy"].ToString(),
                        CreatedOn = row["CreatedOn"].ToString(),
                        TypeName = row["TypeName"].ToString(),
                        Space = row["Space"].ToString(),
                        Description = row["Description"].ToString(),
                        ExtSystem = row["ExtSystem"].ToString(),
                        ExtObject = row["ExtObject"].ToString(),
                        ExtIdentifier = row["ExtIdentifier"].ToString(),
                        SerialNumber = row["SerialNumber"].ToString(),
                        InstallationDate = row["InstallationDate"].ToString(),
                        WarrantyStartDate = row["WarrantyStartDate"].ToString(),
                        TagNumber = row["TagNumber"].ToString(),
                        AssetIdentifier = row["AssetIdentifier"].ToString(),
                        Area = row["Area"].ToString(),
                        Length = row["Length"].ToString()
                    });
                }

                // Systems
                var dtSystem = ReadSheetToDataTable(filePath, "System");
                Systems.Clear();
                string[] systemColumns = new string[]
                {
                    "Name", "CreatedBy", "CreatedOn", "Category", "ComponentNames",
                    "ExtSystem", "ExtObject", "ExtIdentifier", "Description"
                };
                CheckColumnsExist(dtSystem, "System", systemColumns);
                foreach (DataRow row in dtSystem.Rows)
                {
                    Systems.Add(new HeThong
                    {
                        Name = row["Name"].ToString(),
                        CreatedBy = row["CreatedBy"].ToString(),
                        CreatedOn = row["CreatedOn"].ToString(),
                        Category = row["Category"].ToString(),
                        ComponentNames = row["ComponentNames"].ToString(),
                        ExtSystem = row["ExtSystem"].ToString(),
                        ExtObject = row["ExtObject"].ToString(),
                        ExtIdentifier = row["ExtIdentifier"].ToString(),
                        Description = row["Description"].ToString()
                    });
                }
                var dt = ReadSheetToDataTable(filePath, "Attribute"); // Sheet Excel phải tên "ThuocTinh"
                Attributes.Clear();

                string[] requiredColumns = new string[]
                {
                "Name", "CreatedBy", "CreatedOn", "Category", "SheetName", "RowName", "Value", "Unit",
                "ExtSystem", "ExtObject", "ExtIdentifier", "Description", "AllowedValues"
                };

                foreach (var col in requiredColumns)
                {
                    if (!dt.Columns.Contains(col))
                        throw new Exception($"Sheet 'Attribute' thiếu cột '{col}'.");
                }

                foreach (DataRow row in dt.Rows)
                {
                    Attributes.Add(new ThuocTinh
                    {
                        Name = row["Name"].ToString(),
                        CreatedBy = row["CreatedBy"].ToString(),
                        CreatedOn = row["CreatedOn"].ToString(),
                        Category = row["Category"].ToString(),
                        SheetName = row["SheetName"].ToString(),
                        RowName = row["RowName"].ToString(),
                        Value = row["Value"].ToString(),
                        Unit = row["Unit"].ToString(),
                        ExtSystem = row["ExtSystem"].ToString(),
                        ExtObject = row["ExtObject"].ToString(),
                        ExtIdentifier = row["ExtIdentifier"].ToString(),
                        Description = row["Description"].ToString(),
                        AllowedValues = row["AllowedValues"].ToString()
                    });
                }
                var dtCoorodinate = ReadSheetToDataTable(filePath, "Coordinate");
                Coordinates.Clear();

                string[] CoordinateColumns = new string[]
                {
    "Name", "CreatedBy", "CreatedOn", "Category", "SheetName", "RowName",
    "CoordinateXAxis", "CoordinateYAxis", "CoordinateZAxis", "ExtSystem", "ExtObject",
    "ExtIdentifier", "ClockwiseRotation", "ElevationalRotation", "YawRotation"
                };

                foreach (var col in CoordinateColumns)
                {
                    if (!dtCoorodinate.Columns.Contains(col))
                        throw new Exception($"Sheet 'Coordinate' thiếu cột '{col}'.");
                }

                foreach (DataRow row in dtCoorodinate.Rows)
                {
                    Coordinates.Add(new DieuPhoi
                    {
                        Name = row["Name"].ToString(),
                        CreatedBy = row["CreatedBy"].ToString(),
                        CreatedOn = row["CreatedOn"].ToString(),
                        Category = row["Category"].ToString(),
                        SheetName = row["SheetName"].ToString(),
                        RowName = row["RowName"].ToString(),
                        CoordinateXAxis = row["CoordinateXAxis"].ToString(),
                        CoordinateYAxis = row["CoordinateYAxis"].ToString(),
                        CoordinateZAxis = row["CoordinateZAxis"].ToString(),
                        ExtSystem = row["ExtSystem"].ToString(),
                        ExtObject = row["ExtObject"].ToString(),
                        ExtIdentifier = row["ExtIdentifier"].ToString(),
                        ClockwiseRotation = row["ClockwiseRotation"].ToString(),
                        ElevationalRotation = row["ElevationalRotation"].ToString(),
                        YawRotation = row["YawRotation"].ToString(),
                    });
                }
            }

            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message);
                System.Windows.MessageBox.Show($"Lỗi load Excel: {ex.Message}");
            }
        }

        private void CheckColumnsExist(DataTable dt, string sheetName, string[] requiredColumns)
        {
            foreach (var col in requiredColumns)
            {
                if (!dt.Columns.Contains(col))
                    throw new Exception($"Sheet '{sheetName}' thiếu cột '{col}'.");
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
