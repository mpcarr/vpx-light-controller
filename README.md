# VPX Light Controller

A class to help vpx originals use nFozzy Lampz by managing light states and sequences. Supports managing multiple light states per light, RGB lights, Syncing with VPX light sequences, custom key frame light sequences, import / export of light sequences to show creator and several utility functions for interacting with lights.

## Table of Contents

  - [Installing](#installing)
  - [Examples](#examples)
  - [Register Lights](#registerLights)
  - [Managing Lights](#managingLights)
    - [Light On/Off](#lightOn)
    - [Light Blink](#lightBlink)
    - [Light Color](#lightColor)
    - [Light Level](#lightLevel)
    - [Light Utilities](#lightUtilities)
        - [Light On With Color](#lightOnWithColor)
        - [Light On With Flicker](#lightFlicker)
        - [Light Pulse](#lightPulse)
        - [Light Pulse With Profile](#lightPulseWithProfile)
  - [Managing Shots](#managingShots)
    - [Add Shot](#addShot)
    - [Remove Shot](#removeShot)
    - [Remove All Shots](#removeallShots)
    - [Is Shot Lit](#IsShotLit)
  - [Syncing With VPX Lights](#syncWithVPXLights)
  - [Light Sequences](#lightSequences)    
    - [VPX Sequences](#vpxSequences)
    - [Custom Sequences](#customSequences)
    - [MPF Show Creator](#showCreator)
      - [Exporting Lights](#exportingLights)
      - [Importing Light Shows](#importingLightShows)


# Installing

- Download the latest release from github releases and copy the vbs code from **lightController.vbs** into your vpx table script.

- Copy the vbs code from **lampz.vbs** into your table script.

# Examples

TODO

# Register Lights <a name="registerLights"></a>

You need to register the lights you want to control with the class. To do this, add the below code after you have assigned all your lights with Lampz. The best place to do this is at the end of the InitLampNF sub.

```
lightCtrl.RegisterLights
```

# Managing Lights <a name="managingLights"></a>

To control your lights you can call the following functions on the light controller object. All these require you to pass the vpx light as a parameter. E.g. The below examples assume you have a light in vpx called L01

## Light On/Off

To turn a light on you use: **On**.

```
lightCtrl.LightOn L01
```

![Alt Text](https://github.com/mpcarr/vpx-light-controller/blob/main/images/lighton.gif?raw=true)



## Light Blink

To start a light blinking you use: **Blink**. This will use the blink pattern of the vpx light

```
lightCtrl.Blink L01
```

## Light Color

VPX lights have color and fullColor properties. To change the light color use: **Color**

```
lightCtrl.LightColor L01, RGB(255,0,0) 'Sets the color property to Red

lightCtrl.LightColor L01, Array(RGB(255,0,0), Null) 'Sets the color to Red

lightCtrl.LightColor L01, Array(RGB(255,0,0), RGB(0,255,0)) 'Sets the color and fullColor propteries to Red & Green

lightCtrl.LightColor L01, Array(Null, RGB(0,0,255)) 'Sets the fullColor to Blue
```

## Light Level

All of the lights registerd with the controller will be set to modulate brightness between 0-100. Lampz will fade the light upto this level. To change the level use: **Level**

```
lightCtrl.Level L01, 50 'Set the lights maximum brightness to 50%

lightCtrl.Level L01, 100 'Set the lights maximum brightness to 100%
```

## Light Utilities

To help control common light events there are some utility functions available:

### Light On With Color

Turns a light shilst also setting the color

```
lightCtrl.LightOnWithColor L01, RGB(255,0,0)
```

### Light On With Flicker

Turns a light on with a short flickering sequence

```
lightCtrl.LightOnWithFlicker L01
```

### Light Pulse

Pulses a light with the default pulse sequence. The state of the light does not change. E.g. after the pulse has finished, if the light was previously on, it will be on again.

The second parameter is the number of times to repeat the pulse.

```
lightCtrl.Pulse L01, 0 'Pulse the light once, no repeat (total 1 pulse)

lightCtrl.Pulse L01, 1 'Pulse the light, repeat the pulse once (total 2 pulses)
```

### Light Pulse With Profile

Pulses a light with a custom pulse sequence. The state of the light does not change. E.g. after the pulse has finished, if the light was previously on, it will be on again.

The second parameter is the pulse sequence to use. Each element in the array is the brightness of the light for that frame. Each frame is 20ms so the 5 element seq below would last 100ms.

```
lightCtrl.PulseWithProfile L01, Array(10,50,100,50,0) 0
``` 

# Managing Shots

Aside from turning lights on and making them blink, you can add dfferent sequences to lights which run one after the other. You can manually manage this by creating sequence runners and adding / removing sequences. For convience the most common senarnio is to have one light lit for lots of differnt reasons. You may have a light lit for a mode shot but also for a jackpot or combo. The light controller lets you add multiple "Shots" to a light.

## Add Shot

To add a new shot to a light use: **AddShot**

```
lightCtrl.AddShot "Name", L01, RGB(255,0,0)
```

The first parameter is a name for the shot so we can reference it later.
The second parameter is the vpx light object
The third parameter is the color you want this shot to be.

## Remove Shot

To remove a shot from a light use: **RemoveShot**

```
lightCtrl.RemoveShot "Name"
```

Call remove shot with the name of the shot you want to remove

## Remove All Shots

To remove all shots from a light use: **RemoveAllShots**

```
lightCtrl.RemoveAllShots L01
```

Call remove all shots with the vpx light object

## Is Shot Lit

To check if a shot is current active on the light use: **IsShotLit**

```
lightCtrl.IsShotLit "Name", L01
```

The first parameter is the name of the shot
The second parameter is the vpx light object

# Sync With VPX Lights

A simple way to use Lampz in your table without any complex seqeuences is to sync your control lights via the GetInPlayState property. You can use the light controller todo this for you by calling **SyncWithVpxLights**. Once enabled, the controller will update Lampz with your control lights.

```
lightCtrl.SyncWithVpxLights aLights 'alights is the collection you want to sync with. e.g. All Lights. You can create a smaller set of lights to sync if you wish.
```

# Light Sequences

## VPX Sequences

The light controller can sync your lights to VPX light sequences

In the example below you start the vpx sequence **SeqCircleOutOn**. Then call the light controller **SyncWithVpxLights** function with the collection the vpx light sequence uses. This will start syncing your lights.

```
SomeVPXLightSeq.Play SeqCircleOutOn,50,100
lightCtrl.SyncWithVpxLights VpxCollection

'Stop syncing with vpx lights once the sequence is complete
Sub SomeVPXLightSeq_PlayDone()
	lightCtrl.StopSyncWithVpxLights()
End Sub
```

By default this will use the current color of the light. If you want to override the color of all the lights during the sequence you can use **SetVpxSyncLightColor**

```
SomeVPXLightSeq.Play SeqCircleOutOn,50,100
lightCtrl.SyncWithVpxLights VpxCollection
lightCtrl.SetVpxSyncLightColor RGB(255,0,0)
```

## Custom Sequences

The controller supports writing your own custom sequences using a **SequenceRunner** and **LightSequence** items. Each Sequence Runner added to the controller is ran concurrently with other sequence runners. Each Light Sequence added to a runner is ran one after the other (following).

Lets say you have a set of 3 lights and you want to light one after the other on repeat.

The sequence might this:

 - First frame: 
   - Light 1 On
 - Second frame:
   - Light 1 Off
   - Light 2 On
 - Third Frame:
   - Light 2 Off
   - Light 3 Off
 - Fourth Frame: 
   - Light 3 Off

Using the controller you can create a sequence runner for this using **CreateSeqRunner**

```
lightCtrl.CreateSeqRunner "_RunnerName_" 'Runner name can be anything you want
```

Then you need to create the actual sequence using the **LCSeq** class.

```
Dim lSeqTest: Set lSeqTest = new LCSeq
lSeqTest.Name = "lSeqTest"

lSeqTest.Sequence = Array( _
                    Array("L01|100"), _
                    Array("L01|0", "L02|100"), _
                    Array("L02|0", "L03|100"), _
                    Array("L03|0"))

lSeqTest.UpdateInterval = 40
lSeqTest.Color = Null
lSeqTest.Repeat = True

```

Finally, when you want to run the sequence, you need to add it to the runner you created using **AddLightSeq**.

```
lightCtrl.AddLightSeq "_RunnerName_", lSeqTest
```

As the above example is set to repeat, you can stop it by using **RemoveLightSeq**

```
lightCtrl.RemoveLightSeq "_RunnerName_", lSeqTest
```

# Show Creator

For more complex light sequences the controller supports exporting your table lights so that they can be imported a tool called **showcreator**

https://github.com/missionpinball/showcreator

Show Creator is a tool for creating lightshows in MPF (Mission Pinball Framework). It takes a YAML file which describes your light locations and exports a YAML file with key frame data of the light show you created.

We can use some functions on the light controller to help us integrate with MPF Show Creator

## Exporting Lights

You can export the position of your lights by using **CompileLights**. You need to pass in a vpx collection of the lights you want to export (this lets you have more fine grain control over which lights to animate in show creator).

To export lights you need to have the variable **cGameName** set in your table. This is used to export the file to your computer in a subfolder called **cGameName_LightShows**

```
lightCtrl.CompileLights _VPXCollection_, "filename"
```

This will export a file you can use in MPF Show Creator.

## Importing Light Shows

You can import the created lights into your table with the **LoadLightShows** command. This will read light show files saved from Show Creator and compile them to light sequences the controller can run. **LoadLightShows** will also output a file called **lights-out.txt** which contains the compiled vbs code for the light sequence. For development is easier to call **LoadLightShows**, however you should copy the compiled lightshows into your tablescript before sharing it.

```
lightCtrl.LoadLightShows
```