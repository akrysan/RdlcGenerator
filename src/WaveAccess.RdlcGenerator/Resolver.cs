namespace WaveAccess.RdlcGenerator {
    using System;

    class Resolver : IResolver {
        public object Resolve(Type type) {
            return Activator.CreateInstance(type);
        }
    }
}
