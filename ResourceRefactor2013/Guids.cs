// Guids.cs
// MUST match guids.h
using System;

namespace ResourceRefactoring.ResourceRefactor2013
{
    static class GuidList
    {
        public const string guidResourceRefactor2013PkgString = "1bd39d38-1eb0-423f-bc40-46301c752ace";
        public const string guidResourceRefactor2013CmdSetString = "ab43283b-3fe3-43d8-bb35-b2758157469c";

        public static readonly Guid guidResourceRefactor2013CmdSet = new Guid(guidResourceRefactor2013CmdSetString);
    };
}