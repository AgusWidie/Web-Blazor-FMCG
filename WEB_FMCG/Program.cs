using Blazored.LocalStorage;
using Blazored.Toast;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using WEB_FMCG;
using WEB_FMCG.Services;
using OfficeOpenXml;

var builder = WebAssemblyHostBuilder.CreateDefault(args);

//ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
ExcelPackage.License.SetNonCommercialOrganization("Logic Soft Computer");
builder.RootComponents.Add<App>("#app");
builder.RootComponents.Add<HeadOutlet>("head::after");

builder.Services.AddScoped(sp =>
    new HttpClient { BaseAddress = new Uri("https://api.coresight-hub.cloud/") });

builder.Services.AddSingleton<AuthService>();
builder.Services.AddBlazoredLocalStorage();
builder.Services.AddBlazoredToast();

builder.Services.AddScoped<ILoginService, LoginService>();
builder.Services.AddScoped<ILogoutService, LogoutService>();
builder.Services.AddScoped<IAreaService, AreaService>();
builder.Services.AddScoped<IAttendanceLeaveService, AttendanceLeaveService>();
builder.Services.AddScoped<IAttendanceOverTimeService, AttendanceOverTimeService>();
builder.Services.AddScoped<IAttendancePermitService, AttendancePermitService>();
builder.Services.AddScoped<IAttendanceService, AttendanceService>();
builder.Services.AddScoped<ICityService, CityService>();
builder.Services.AddScoped<IMechanismeService, MechanismeService>();
builder.Services.AddScoped<ILoginDeviceService, LoginDeviceService>();
builder.Services.AddScoped<IMenuService, MenuService>();
builder.Services.AddScoped<IOrderProductService, OrderProductService>();
builder.Services.AddScoped<IPhotoProductCOCCashierService, PhotoProductCOCCashierService>();
builder.Services.AddScoped<IPhotoProductHomeShelfService, PhotoProductHomeShelfService>();
builder.Services.AddScoped<IPhotoProductPriceCompetitorService, PhotoProductPriceCompetitorService>();
builder.Services.AddScoped<IPhotoProductSecondaryDisplayService, PhotoProductSecondaryDisplayService>();
builder.Services.AddScoped<IPhotoPromotionService, PhotoPromotionService>();
builder.Services.AddScoped<IProductCategoryService, ProductCategoryService>();
builder.Services.AddScoped<IProductPricingService, ProductPricingService>();
builder.Services.AddScoped<IProductService, ProductService>();
builder.Services.AddScoped<IProductVariantService, ProductVariantService>();
builder.Services.AddScoped<IProductStockService, ProductStockService>();
builder.Services.AddScoped<IRegionService, RegionService>();
builder.Services.AddScoped<IRegionStoreService, RegionStoreService>();
builder.Services.AddScoped<IRentalDisplayService, RentalDisplayService>();
builder.Services.AddScoped<IRoleService, RoleService>();
builder.Services.AddScoped<ISurveyService, SurveyService>();
builder.Services.AddScoped<ISalesService, SalesService>();
builder.Services.AddScoped<IStoreService, StoreService>();
builder.Services.AddScoped<IStoreVisitService, StoreVisitService>();
builder.Services.AddScoped<ISubmissionLeaveService, SubmissionLeaveService>();
builder.Services.AddScoped<IPositionService, PositionService>();
builder.Services.AddScoped<IDepartmentService, DepartmentService>();
builder.Services.AddScoped<IDivisionService, DivisionService>();
builder.Services.AddScoped<ILogStockService, LogStockService>();
builder.Services.AddScoped<ILogSalesTrackingService, LogSalesTrackingService>();
builder.Services.AddScoped<ISalesTrackingService, SalesTrackingService>();
builder.Services.AddScoped<ISurveyQuestionService, SurveyQuestionService>();
builder.Services.AddScoped<ISurveyResponseService, SurveyResponseService>();

await builder.Build().RunAsync();

