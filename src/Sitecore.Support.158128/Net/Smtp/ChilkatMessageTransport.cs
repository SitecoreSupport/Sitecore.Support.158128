// © 2016 Sitecore Corporation A/S. All rights reserved.

using System;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Chilkat;
using Sitecore.Diagnostics;
using Sitecore.EDS.Core.Dispatch;
using Sitecore.EDS.Core.Exceptions;
using Sitecore.EDS.Core.Net.Smtp;
using Sitecore.Modules.EmailCampaign.Factories;
using Debug = System.Diagnostics.Debug;
using Task = System.Threading.Tasks.Task;

namespace Sitecore.Support.EDS.Core.Net.Smtp
{
  /// <summary>
  /// Defines chilkat message transport envelope.
  /// </summary>
  public class ChilkatMessageTransport : MessageTransport<ITransportClient>
  {
    private readonly int _maxTries;
    private readonly int _delay;
    /// <summary>
    /// Initializes a new instance of the <see cref="ChilkatMessageTransport" /> class.
    /// </summary>
    /// <param name="message">The email message.</param>
    public ChilkatMessageTransport(EmailMessage message, int maxTries, int delay)
        : base(message)
    {
      Email = ParseMessage(message);
      _maxTries = maxTries;
      _delay = delay;
    }

    /// <summary>
    /// Gets the email message size.
    /// </summary>
    public override int Size
    {
      get { return Email.Size; }
    }

    /// <summary>
    /// Gets the email.
    /// </summary>
    /// <value>
    /// The email.
    /// </value>
    [NotNull]
    protected Email Email { get; private set; }

    /// <summary>
    /// Sends the transport message asynchronously.
    /// </summary>
    /// <param name="client">The client.</param>
    /// <returns>
    /// The result of message dispatch operation <see cref="DispatchResult" />.
    /// </returns>
    public override Task<DispatchResult> SendAsync(ITransportClient client)
    {
      Assert.ArgumentNotNull(client, "client");
      return SendTaskAsync(client);
    }

    /// <summary>
    /// Sends the transport message asynchronously.
    /// </summary>
    /// <param name="client">The client.</param>
    /// <returns>
    /// The result of message dispatch operation <see cref="DispatchResult" />.
    /// </returns>
    private async Task<DispatchResult> SendTaskAsync(ITransportClient client)
    {
      Assert.ArgumentNotNull(client, "client");

      var stopWatch = Stopwatch.StartNew();
      await RetryOnFault(() => client.SendAsync(Email), _maxTries);
      stopWatch.Stop();
      var endTime = stopWatch.ElapsedMilliseconds;

      var dispatchResult = new DispatchResult();
      dispatchResult.Statistics.Add("SendingTime", endTime.ToString());
      dispatchResult.Statistics.Add("Size", Size.ToString(CultureInfo.InvariantCulture));
      return dispatchResult;
    }

    /// <summary>
    /// Maps to chilkat email type.
    /// </summary>
    /// <param name="message">The message.</param>
    /// <returns>The email object for chilkat</returns>
    private static Email ParseMessage(EmailMessage message)
    {
      var parsedEmail = new Email
      {
        Subject = message.Subject,
        FromAddress = message.FromAddress,
        FromName = message.FromName,
        Charset = message.Charset,
        Sender = message.ReturnPath,
        BounceAddress = message.ReturnPath
      };

      if (message.ContentType == MessageContentType.Html)
      {
        parsedEmail.SetHtmlBody(message.HtmlBody);
        if (!string.IsNullOrEmpty(message.PlainTextBody))
        {
          parsedEmail.AddPlainTextAlternativeBody(message.PlainTextBody);
        }
      }
      else
      {
        parsedEmail.Body = message.PlainTextBody;
      }

      parsedEmail.AddHeaderField("X-Priority", ((int)message.Priority).ToString(CultureInfo.InvariantCulture));

      foreach (string header in message.Headers)
      {
        parsedEmail.AddHeaderField(header, message.Headers[header]);
      }

      foreach (var recipient in message.Recipients)
      {
        parsedEmail.AddTo(string.Empty, recipient);
      }

      foreach (var attachment in message.Attachments.Where(a => !a.IsEmbedded))
      {
        parsedEmail.AddDataAttachment(attachment.Name, attachment.ByteContent);
      }

      foreach (var embeddeData in message.Attachments.Where(a => a.IsEmbedded && !string.IsNullOrEmpty(a.ContentId)))
      {
        var contentId = parsedEmail.AddRelatedData(embeddeData.Name, embeddeData.ByteContent);
        parsedEmail.SetReplacePattern(embeddeData.ContentId, "cid:" + contentId);
      }

      return parsedEmail;
    }

    /// <summary>
    /// The retry on fault.
    /// </summary>
    /// <param name="function">The function.</param>
    /// <param name="maxTries">The max tries.</param>
    /// <returns>
    /// The <see cref="Task" />.
    /// </returns>
    private async Task RetryOnFault(Func<Task> function, int maxTries)
    {
      for (var i = 0; i < maxTries; i++)
      {
        try
        {
          await function().ConfigureAwait(false);
          break;
        }
        // fix 158128
        catch (TransportException)
        {

          if (i == maxTries - 1)
          {
            throw;
          }
          EcmFactory.GetDefaultFactory().Io.Logger.LogInfo("Sending operation is failed. Connection attempt: " + (i + 1) + ". It is retried.");
          Thread.Sleep(_delay);
        }
      }
    }
  }
}
