using System;
using System.Windows;
using System.Windows.Input;
using System.Collections.Generic;
using UserActivityTracker.FileFormat;

namespace UserActivityTracker
{
    /// <summary>
    /// Provides the ability to record a series of user actions on a specified <see cref="FrameworkElement"/>.
    /// </summary>
    public class Recorder
    {
        /// <summary>
        /// Gets the <see cref="FrameworkElement"/> that is set to be recorded.
        /// </summary>
        public FrameworkElement Element { get; }

        /// <summary>
        /// Indicates whether the recording has started yet.
        /// </summary>
        public bool IsRecording
        {
            get
            {
                return _isRecording;
            }
        }
        private bool _isRecording;

        /// <summary>
        /// Gets or sets the number of basic user actions that are recorded per second, including moving the mouse. The default value is 15.
        /// </summary>
        public int FrameRate
        {
            get
            {
                return _frameRate;
            }
            set
            {
                _frameRate = value;
                _frameDuration = 1000 / value;
            }
        }
        private int _frameRate;
        private int _frameDuration;

        /// <summary>
        /// Gets or sets whether to optimize recording data based on <see cref="FrameRate"/>. The default value is <see langword="true"/>.
        /// </summary>
        public bool OptimizeData { get; set; }

        /// <summary>
        /// Gets or sets whether to record mouse actions from the user. The default value is <see langword="true"/>.
        /// </summary>
        public bool RecordMouseActions { get; set; }

        /// <summary>
        /// Gets or sets whether to record keyboard actions from the user. The default value is <see langword="true"/>.
        /// </summary>
        public bool RecordKeyboardActions { get; set; }

        private Structure session;
        private UserAction lastAction;
        private int lastActionTime;

        /// <summary>
        /// Initializes a new instance of the <see cref="Recorder"/> class on a specified <see cref="FrameworkElement"/>.
        /// </summary>
        /// <param name="element">The <see cref="FrameworkElement"/> that is set to be recorded.</param>
        public Recorder(FrameworkElement element)
        {
            this.Element = element;
            this.FrameRate = 15;
            this.OptimizeData = true;
            this.RecordMouseActions = true;
            this.RecordKeyboardActions = true;

            _isRecording = false;
        }

        /// <summary>
        /// Starts to record user actions on <see cref="Element"/> if the recording has not started yet.
        /// </summary>
        /// <param name="startingConfig">An optional <see langword="string"/> value to store customized starting configurations. Cannot include the character ";" in it.</param>
        /// <returns><see langword="true"/> if the recording was started successfully; otherwise, <see langword="false"/>.</returns>
        public bool Start(string startingConfig = "")
        {
            if (this.IsRecording || this.Element == null || this.FrameRate <= 0
                || startingConfig.Contains(";"))
            {
                return false;
            }
            else
            {
                _isRecording = true;
            }

            this.Element.Focus();

            session = new Structure
            {
                FrameRate = this.FrameRate,
                StartingWidth = this.Element.ActualWidth,
                StartingHeight = this.Element.ActualHeight,
                StartingConfig = startingConfig,
                Actions = new List<string>()
            };
            lastAction = new UserAction
            {
                ActionType = UserActionType.Unknown
            };
            lastActionTime = Environment.TickCount;

            this.Element.SizeChanged += AddSizeChanged;
            this.Element.PreviewMouseMove += AddMouseMove;
            this.Element.PreviewMouseDown += AddMouseDown;
            this.Element.PreviewMouseUp += AddMouseUp;
            this.Element.PreviewMouseWheel += AddMouseWheel;
            this.Element.PreviewKeyDown += AddKeyDown;
            this.Element.PreviewKeyUp += AddKeyUp;

            return true;
        }

        /// <summary>
        /// Stops the current recording on <see cref="Element"/> if the recording is ongoing.
        /// </summary>
        /// <returns><see langword="true"/> if the recording was stopped successfully; otherwise, <see langword="false"/>.</returns>
        public bool Stop()
        {
            if (!this.IsRecording || this.Element == null)
            {
                return false;
            }
            else
            {
                _isRecording = false;
            }

            this.Element.SizeChanged -= AddSizeChanged;
            this.Element.PreviewMouseMove -= AddMouseMove;
            this.Element.PreviewMouseDown -= AddMouseDown;
            this.Element.PreviewMouseUp -= AddMouseUp;
            this.Element.PreviewMouseWheel -= AddMouseWheel;
            this.Element.PreviewKeyDown -= AddKeyDown;
            this.Element.PreviewKeyUp -= AddKeyUp;

            return true;
        }

        /// <summary>
        /// Saves the current recording on <see cref="Element"/> if the recording exists.
        /// </summary>
        /// <returns>A <see langword="string"/> representation of all user actions on <see cref="Element"/> during the recording.</returns>
        public string Save()
        {
            if (session == null)
            {
                return "";
            }

            return Structure.Serialize(session);
        }

        /// <summary>
        /// Adds a custom <see langword="string"/> message to the recording that can be outputted to the logs.
        /// </summary>
        /// <param name="message">A custom <see langword="string"/> that will be added to the recording. Cannot include the characters ";" and "'" in it.</param>
        /// <returns><see langword="true"/> if the message was added successfully; otherwise, <see langword="false"/>.</returns>
        public bool LogMessage(string message)
        {
            if (message.Contains(";") || message.Contains("'"))
            {
                return false;
            }

            AddUserAction(UserActionType.Message, false, $"\'{message}\'");

            return true;
        }

        private void AddUserAction(UserActionType actionType, bool addPause, params object[] actionParameters)
        {
            if (!this.IsRecording)
            {
                return;
            }

            if (addPause)
            {
                int currentTime = Environment.TickCount;
                int extraTime = currentTime - lastActionTime - _frameDuration;
                if (extraTime > 0)
                {
                    lastAction = new UserAction()
                    {
                        ActionType = UserActionType.Pause,
                        ActionParameters = new object[] { extraTime }
                    };
                    session.Actions.Add(lastAction.ToString());
                }
                lastActionTime = currentTime;
            }

            lastAction = new UserAction()
            {
                ActionType = actionType,
                ActionParameters = actionParameters
            };
            session.Actions.Add(lastAction.ToString());
        }

        private void AddSizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (this.OptimizeData && lastAction.ActionType == UserActionType.Resize
                && Environment.TickCount - lastActionTime < _frameDuration)
            {
                session.Actions.RemoveAt(session.Actions.Count - 1);
            }

            AddUserAction(UserActionType.Resize, true, this.Element.ActualWidth, this.Element.ActualHeight);
        }

        private void AddMouseMove(object sender, MouseEventArgs e)
        {
            if (!this.RecordMouseActions || Environment.TickCount - lastActionTime < _frameDuration)
            {
                return;
            }

            Point position = e.GetPosition(this.Element);
            AddUserAction(UserActionType.MouseMove, true, position.X, position.Y);
        }

        private void AddMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (!this.RecordMouseActions)
            {
                return;
            }

            Point position = e.GetPosition(this.Element);
            AddUserAction(UserActionType.MouseDown, true, position.X, position.Y, (int)e.ChangedButton);
        }

        private void AddMouseUp(object sender, MouseButtonEventArgs e)
        {
            if (!this.RecordMouseActions)
            {
                return;
            }

            Point position = e.GetPosition(this.Element);
            AddUserAction(UserActionType.MouseUp, true, position.X, position.Y, (int)e.ChangedButton);
        }

        private void AddMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (!this.RecordMouseActions)
            {
                return;
            }

            Point position = e.GetPosition(this.Element);
            int delta = e.Delta;

            if (this.OptimizeData && lastAction.ActionType == UserActionType.MouseWheel
                && Environment.TickCount - lastActionTime < _frameDuration
                && lastAction.ActionParameters.Length >= 3
                && double.TryParse(lastAction.ActionParameters[0].ToString(), out double lastX)
                && double.TryParse(lastAction.ActionParameters[1].ToString(), out double lastY)
                && position.X == lastX && position.Y == lastY
                && int.TryParse(lastAction.ActionParameters[2].ToString(), out int lastDelta))
            {
                delta += lastDelta;
                session.Actions.RemoveAt(session.Actions.Count - 1);
            }

            AddUserAction(UserActionType.MouseWheel, true, position.X, position.Y, delta);
        }

        private void AddKeyDown(object sender, KeyEventArgs e)
        {
            if (!this.RecordKeyboardActions)
            {
                return;
            }

            AddUserAction(UserActionType.KeyDown, true, KeyInterop.VirtualKeyFromKey(e.Key));
        }

        private void AddKeyUp(object sender, KeyEventArgs e)
        {
            if (!this.RecordKeyboardActions)
            {
                return;
            }

            AddUserAction(UserActionType.KeyUp, true, KeyInterop.VirtualKeyFromKey(e.Key));
        }
    }
}
