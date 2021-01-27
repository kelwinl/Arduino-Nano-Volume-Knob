Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;
[Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IAudioEndpointVolume
{
    // f(), g(), ... are unused COM method slots. Define these if you care
    int f(); int g(); int h(); int i();
    int SetMasterVolumeLevelScalar(float fLevel, System.Guid pguidEventContext);
    int j();
    int GetMasterVolumeLevelScalar(out float pfLevel);
    int k(); int l(); int m(); int n();
    int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, System.Guid pguidEventContext);
    int GetMute(out bool pbMute);
}
[Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDevice
{
    int Activate(ref System.Guid id, int clsCtx, int activationParams, out IAudioEndpointVolume aev);
}
[Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDeviceEnumerator
{
    int f(); // Unused
    int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice endpoint);
}
[ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")] class MMDeviceEnumeratorComObject { }
public class Audio
{
    static IAudioEndpointVolume Vol()
    {
        var enumerator = new MMDeviceEnumeratorComObject() as IMMDeviceEnumerator;
        IMMDevice dev = null;
        Marshal.ThrowExceptionForHR(enumerator.GetDefaultAudioEndpoint(/*eRender*/ 0, /*eMultimedia*/ 1, out dev));
        IAudioEndpointVolume epv = null;
        var epvid = typeof(IAudioEndpointVolume).GUID;
        Marshal.ThrowExceptionForHR(dev.Activate(ref epvid, /*CLSCTX_ALL*/ 23, 0, out epv));
        return epv;
    }
    public static float Volume
    {
        get { float v = -1; Marshal.ThrowExceptionForHR(Vol().GetMasterVolumeLevelScalar(out v)); return v; }
        set { Marshal.ThrowExceptionForHR(Vol().SetMasterVolumeLevelScalar(value, System.Guid.Empty)); }
    }
    public static bool Mute
    {
        get { bool mute; Marshal.ThrowExceptionForHR(Vol().GetMute(out mute)); return mute; }
        set { Marshal.ThrowExceptionForHR(Vol().SetMute(value, System.Guid.Empty)); }
    }
}
'@


$port= new-Object System.IO.Ports.SerialPort COM4,9600,None,8,one
try
    {
        $port.Open()
        $encoderValue = $port.ReadLine()
        $encoderPrev = $port.ReadLine()
    }
catch
    {
        $port= new-Object System.IO.Ports.SerialPort COM4,9600,None,8,one
        try
            {
                $port.Open()
                $encoderValue = $port.ReadLine()
                $encoderPrev = $port.ReadLine()
            }
        catch{Start-Sleep -Second 2}
    }




$fvolume = [audio]::Volume

while($true)
    {
        try{$encoderValue = $port.ReadLine()}
        catch
            {
                $port= new-Object System.IO.Ports.SerialPort COM4,9600,None,8,one
                try{$port.Open()}
                catch{Start-Sleep -Second 2}
            }



        if($encoderValue -ge 999)
            {
                if([audio]::Mute -eq $true)
                    {
                        [audio]::Mute = $false
                    }
                else
                    {
                        [audio]::Mute = $true
                    }
                $encoderValue = $encoderPrev
            }
        elseif($encoderPrev -ne $encoderValue)
            {
                [audio]::Mute = $false
                if($encoderValue -gt $encoderPrev)
                    {
                        if($fvolume -le 0.8)
                            {
                                [audio]::Volume  = $fvolume + 0.02
                            }
                        else
                            {
                                [audio]::Volume = 1
                            }
                    }
                elseif($encoderValue -lt $encoderPrev)
                    {
                        if($fvolume -ge 0.02)
                            {
                                [audio]::Volume  = $fvolume - 0.02
                            }
                        else
                            {
                                [audio]::Volume = 0
                            }
                    }
                Write-Output $encoderValue
                Write-Output $encoderPrev
                Write-Output $fvolume
            }
        
        $encoderPrev = $encoderValue
        $fvolume = [audio]::Volume
    }
