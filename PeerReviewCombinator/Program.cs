using Microsoft.Extensions.Configuration;
using PeerReviewCombinator;
using Serilog;

Log.Logger = new LoggerConfiguration()
    .MinimumLevel.Information()
    .WriteTo.Console()
    .CreateLogger();

var config = new ConfigurationBuilder()
    .AddJsonFile("appsettings.json")
    .Build();

(new Combinator(config.GetRequiredSection("Settings").Get<Settings>())).run();
