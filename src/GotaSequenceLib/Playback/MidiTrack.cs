using System;
using Sanford.Multimedia.Midi;

namespace GotaSequenceLib.Playback
{
    public class MidiTrack : AbstractTrack
    {
        private readonly Sanford.Multimedia.Midi.Track _track = new Sanford.Multimedia.Midi.Track();

        // *SEQ properties
        private bool _tie = false;
        override public bool Tie
        {
            set
            {
                _tie = value;
                Message(new ChannelMessage(ChannelCommand.Controller, Index, (int)ControllerType.AllNotesOff));
            }
        }

        private bool _noteWait = true;
        override public bool NoteWait
        {
            set { _noteWait = value; }
        }

        private bool _portamento;
        override public bool Portamento
        {
            set { _portamento = value; }
        }

        private int _voice;
        override public int Voice
        {
            set { _voice = value; }
        }

        private byte _priority;
        override public byte Priority
        {
            set { _priority = value; }
        }

        private byte _volume;
        override public byte Volume
        {
            set { _volume = value; }
        }

        private byte _expression;
        override public byte Expression
        {
            set { _expression = value; }
        }

        private byte _lfoRange;
        override public byte LFORange
        {
            set { _lfoRange = value; }
        }

        private byte _pitchBendRange;
        override public byte PitchBendRange
        {
            set { _pitchBendRange = value; }
        }

        private byte _lfoSpeed;
        override public byte LFOSpeed
        {
            set { _lfoSpeed = value; }
        }

        private byte _lfoDepth;
        override public byte LFODepth
        {
            set { _lfoDepth = value; }
        }

        private ushort _lfoDelay;
        override public ushort LFODelay
        {
            set { _lfoDelay = value; }
        }

        private ushort _lfoPhase;
        override public ushort LFOPhase
        {
            set { _lfoPhase = value; }
        }

        private ushort _lfoDelayCount;
        override public ushort LFODelayCount
        {
            set { _lfoDelayCount = value; }
        }

        private LFOType _lfoType;
        override public LFOType LFOType
        {
            set { _lfoType = value; }
        }

        private sbyte _pitchBend;
        override public sbyte PitchBend
        {
            set { _pitchBend = value; }
        }

        private sbyte _panpot;
        override public sbyte Panpot
        {
            set { _panpot = value; }
        }

        private sbyte _transpose;
        override public sbyte Transpose
        {
            set { _transpose = value; }
        }

        private byte _attack;
        override public byte Attack
        {
            set { _attack = value; }
        }

        private byte _decay;
        override public byte Decay
        {
            set { _decay = value; }
        }

        private byte _sustain;
        override public byte Sustain
        {
            set { _sustain = value; }
        }

        private byte _hold;
        override public byte Hold
        {
            set { _hold = value; }
        }

        private byte _release;
        override public byte Release
        {
            set { _release = value; }
        }

        private byte _portamentoKey;
        override public byte PortamentoKey
        {
            set { _portamentoKey = value; }
        }

        private byte _portamentoTime;
        override public byte PortamentoTime
        {
            set { _portamentoTime = value; }
        }

        private short _sweepPitch;
        override public short SweepPitch
        {
            set { _sweepPitch = value; }
        }

        private int _bankNum;
        override public int BankNum
        {
            set { _bankNum = value; }
        }

        private void Message(ChannelMessage message)
        {
            _track.Insert((int)_player.ElapsedTicks, message);
        }

        public MidiTrack(byte idx, Player player) : base(idx, player)
        {
        }
    }
}

