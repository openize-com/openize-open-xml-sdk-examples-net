using Openize.Slides;
using Openize.Slides.Common;
using Openize.Slides.Common.Enumerations;
using System;

namespace Openize.OpenXML_SDK.Examples.PowerPoint
{
    /// <summary>
    /// Provides C# code examples to demonstrate different animation effects
    /// using the <a href="https://www.nuget.org/packages/Openize.Slides">Openize.Slides</a> library.
    /// </summary>
    public class AnimationExamples
    {
        private const string presentationPath = "../../../Presentations/Existing/existing.pptx";

        /// <summary>
        /// Applies Zoom animation to a rectangle shape and adds it to a new slide.
        /// </summary>
        public void ApplyZoomAnimation()
        {
            ApplyAnimation(AnimationType.Zoom);
        }

        /// <summary>
        /// Applies FlyIn animation to a rectangle shape and adds it to a new slide.
        /// </summary>
        public void ApplyFlyInAnimation()
        {
            ApplyAnimation(AnimationType.FlyIn);
        }

        /// <summary>
        /// Applies Spin animation to a rectangle shape and adds it to a new slide.
        /// </summary>
        public void ApplySpinAnimation()
        {
            ApplyAnimation(AnimationType.Spin);
        }

        /// <summary>
        /// Applies FloatIn animation to a rectangle shape and adds it to a new slide.
        /// </summary>
        public void ApplyFloatInAnimation()
        {
            ApplyAnimation(AnimationType.FloatIn);
        }

        /// <summary>
        /// Applies Bounce animation to a rectangle shape and adds it to a new slide.
        /// </summary>
        public void ApplyBounceAnimation()
        {
            ApplyAnimation(AnimationType.Bounce);
        }

        /// <summary>
        /// Generic method to create a rectangle with specified animation and append it to a new slide.
        /// </summary>
        /// <param name="animation">Animation type to apply to the rectangle.</param>
        private void ApplyAnimation(AnimationType animation)
        {
            try
            {
                // Open the existing presentation
                Presentation presentation = Presentation.Open(presentationPath);

                // Create a new slide
                Slide slide = new Slide();

                // Create a rectangle shape
                Rectangle rectangle = new Rectangle
                {
                    Width = 300.0,
                    Height = 300.0,
                    X = 300.0,
                    Y = 300.0,
                    Animation = animation
                };

                // Draw the rectangle on the slide
                slide.DrawRectangle(rectangle);

                // Append the new slide to the presentation
                presentation.AppendSlide(slide);

                // Save the updated presentation
                presentation.Save();
            }
            catch (Exception ex)
            {
                throw new OpenizeException($"Failed to apply {animation} animation.", ex);
            }
        }
    }
}
