//Copyright(c) Microsoft Corporation.All rights reserved.
//Licensed under the MIT License.

using System;

namespace DSnA.WebJob.DocumentParser
{
    public class UnableToDeleteFileException : Exception
    {
        public UnableToDeleteFileException()
        {
        }

        public UnableToDeleteFileException(string message)
            : base(message)
        {
        }

        public UnableToDeleteFileException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }

    public class LoggerException : Exception
    {
        public LoggerException()
        {
        }

        public LoggerException(string message)
            : base(message)
        {
        }

        public LoggerException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}
