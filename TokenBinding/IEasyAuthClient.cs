using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static TokenBinding.EasyAuthTokenClient;

namespace TokenBinding
{
    public interface IEasyAuthClient
    {
        Task<EasyAuthTokenStoreEntry> GetTokenStoreEntry(TokenAttribute attribute);

        Task RefreshToken(TokenAttribute attribute);
    }
}
