// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Threading;

namespace NcTalkOutlookAddIn.Utilities
{
    internal static class ParallelExecution
    {
        internal static void RethrowFirstFailure(
            AggregateException exception,
            CancellationToken cancellationToken)
        {
            if (exception == null)
            {
                throw new ArgumentNullException("exception");
            }

            Exception failure = exception
                .Flatten()
                .InnerExceptions
                .FirstOrDefault(
                    item => !(item is OperationCanceledException));
            if (failure == null)
            {
                cancellationToken.ThrowIfCancellationRequested();
                throw new OperationCanceledException(cancellationToken);
            }

            ExceptionDispatchInfo.Capture(failure).Throw();
        }
    }
}
