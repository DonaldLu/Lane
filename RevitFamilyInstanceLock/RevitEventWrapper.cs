using Autodesk.Revit.UI;
using System;

namespace RevitFamilyInstanceLock
{
    public abstract class RevitEventWrapper<T> : IExternalEventHandler
    {
        private object @lock;

        private T savedArgs;

        private ExternalEvent revitEvent;

        public RevitEventWrapper()
        {
            this.revitEvent = ExternalEvent.Create(this);
            this.@lock = new object();
        }

        public void Execute(UIApplication app)
        {
            object obj = this.@lock;
            T args;
            lock (obj)
            {
                args = this.savedArgs;
                this.savedArgs = default(T);
            }
            this.Execute(app, args);
        }

        public string GetName()
        {
            return base.GetType().Name;
        }

        public void Raise(T args)
        {
            object obj = this.@lock;
            lock (obj)
            {
                this.savedArgs = args;
            }
            this.revitEvent.Raise();
        }

        public abstract void Execute(UIApplication app, T args);
    }
}
