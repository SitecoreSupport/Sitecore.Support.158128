// © 2016 Sitecore Corporation A/S. All rights reserved.

using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using Sitecore.Diagnostics;
using Sitecore.EDS.Core.Dispatch;
using Sitecore.EDS.Core.Exceptions;
using Sitecore.EDS.Core.Net.Smtp;
using Sitecore.EDS.Core.Reporting;
using Sitecore.EDS.Providers.SparkPost.Configuration;
using Sitecore.EDS.Providers.SparkPost.Dispatch;
using Sitecore.Modules.EmailCampaign.Factories;
using Sitecore.StringExtensions;

namespace Sitecore.Support.EDS.Providers.SparkPost.Dispatch
{
  /// <summary>
  /// Defines the DispatchProvider type.
  /// </summary>
  public class DispatchProvider : Sitecore.EDS.Providers.SparkPost.Dispatch.DispatchProvider
  {
    private readonly ConnectionPoolManager _connectionPoolManager;
    private readonly IConfigurationStore _configurationStore;
    private readonly string _returnPath;
    private readonly int _maxTries;
    private readonly int _delay;

    /// <summary>
    /// Initializes a new instance of the <see cref="DispatchProvider" /> class.
    /// </summary>
    /// <param name="connectionPoolManager">The connection pool manager.</param>
    /// <param name="environmentIdentifier">The environment identifier</param>
    /// <param name="configurationStore">The configuration store</param>
    /// <param name="returnPath">Sets the return path address</param>
    /// <param name="maxTries">Sets the max tries</param>
    /// <param name="delay">Sets the delay</param>
    /// 
      // fix 304755
    public DispatchProvider([NotNull] ConnectionPoolManager connectionPoolManager, [NotNull] IEnvironmentId environmentIdentifier, [NotNull] IConfigurationStore configurationStore, [NotNull] string returnPath, [NotNull]  string maxTries = "3", [NotNull]  string delay = "1000")
        : base(connectionPoolManager, environmentIdentifier, configurationStore, returnPath)
    {
      Assert.ArgumentNotNull(connectionPoolManager, "connectionPoolManager");

      _connectionPoolManager = connectionPoolManager;
      _configurationStore = configurationStore;
      _returnPath = returnPath;
      _maxTries = Int32.Parse(maxTries);
      _delay = Int32.Parse(delay);
    }

    /// <summary>
    /// Validates that the SMTP connection can be established.
    /// </summary>
    /// <returns>
    /// <c>true</c> if the SMTP connection can be established.
    /// </returns>
    public override async Task<bool> ValidateDispatchAsync()
    {
      // fix 304754
      for (var i = 0; i < _maxTries; i++)
      {

        var client = await _connectionPoolManager.GetSmtpConnectionAsync();

        // fix 158128
        if (!client.ValidateSmtpConnection().Result)
        {

          if (i == _maxTries - 1)
          {
            return false;
          }
          EcmFactory.GetDefaultFactory().Io.Logger.LogInfo("Connection validation is failed. Connection attempt: " + (i + 1) + ". It is retried.");
          Thread.Sleep(_delay);
        }
        else
        {
          return true;
        }
      }
      return false;
    }

    /// <summary>
    /// Sends the specified message.
    /// </summary>
    /// <param name="message">The message.</param>
    /// <returns>
    /// The <see cref="DispatchResult"/>
    /// </returns>
    protected override async Task<DispatchResult> SendEmailAsync(EmailMessage message)
    {
      for (var i = 0; i < _maxTries; i++)
      {
        message.ReturnPath = _returnPath;
        try
        {
          var chilkatMessageTransport = new ChilkatMessageTransport(message);
          var client = await _connectionPoolManager.GetSmtpConnectionAsync();

          return await chilkatMessageTransport.SendAsync(client);
        }
        // fix 158128
        catch (TransportException)
        {

          if (i == _maxTries - 1)
          {
            throw;
          }
          EcmFactory.GetDefaultFactory().Io.Logger.LogInfo("Sending operation is failed. Connection attempt: " + (i + 1) + ". It is retried.");
          Thread.Sleep(_delay);
        }
      }
      return null;
    }
  }
}