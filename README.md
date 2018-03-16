# In Proc COM Activator

The In Proc COM Activator acts as a proxy for instantiating .NET COM-visible objects.  Normally, `mscoree.dll` is used in the `InProcServer32` registry entry as a proxy for instantiating COM-visible objects.  However, `mscoree.dll` instantiates all .NET objects in the same process into the same `AppDomain`.  This can cause a clash between `AppDomain` properties and events, such as the `AssemblyResolve` event.

## Required Knowledge ##

You should already be familiar with implementing an in-proc COM object.

## Usage ##

The `InProcServer32` registry key under your registered COM object's CLSID is normally populated with the location of `mscoree.dll` on your computer.  This registry value must be changed to point to the InProcComActivator.dll file instead.

The newly created `AppDomain`'s `FriendlyName` will be set to the value of the COM object's CLSID.  If you wish to override this behavior, create an `AppDomain` registry key under `InProcServer32` with the desired name of the AppDomain.
