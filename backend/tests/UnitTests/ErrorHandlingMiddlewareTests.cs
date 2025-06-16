using System;
using System.IO;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using NUnit.Framework;
using MyApp.Middleware;

namespace MyApp.Tests.UnitTests
{
    [TestFixture]
    public class ErrorHandlingMiddlewareTests
    {
        [Test]
        public async Task Invoke_WhenExceptionIsThrown_ReturnsInternalServerErrorAndJsonErrorMessage()
        {
            // Arrange
            var errorMessage = "Test exception";
            RequestDelegate next = (HttpContext context) => throw new Exception(errorMessage);
            var middleware = new ErrorHandlingMiddleware(next);
            var context = new DefaultHttpContext();
            var responseBody = new MemoryStream();
            context.Response.Body = responseBody;

            // Act
            await middleware.Invoke(context);

            // Assert
            Assert.AreEqual((int)HttpStatusCode.InternalServerError, context.Response.StatusCode);
            responseBody.Seek(0, SeekOrigin.Begin);
            var reader = new StreamReader(responseBody);
            var responseText = await reader.ReadToEndAsync();
            StringAssert.Contains(errorMessage, responseText);
            StringAssert.Contains("error", responseText);
        }
    }
}
