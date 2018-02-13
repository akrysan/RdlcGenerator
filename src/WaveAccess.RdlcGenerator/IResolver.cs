namespace WaveAccess.RdlcGenerator {
    using System;

    public interface IResolver {
        object Resolve(Type type);
    }
}
