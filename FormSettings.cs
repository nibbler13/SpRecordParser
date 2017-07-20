using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SpRecordParser {
	public partial class FormSettings : Form {
		private TimeSpan workingTimeBegin = new TimeSpan(8, 0, 0);
		private TimeSpan workingTimeEnd = new TimeSpan(21, 0, 0);
		private bool ignoreInternalMissedCalls = true;
		private bool ignoreNonworkingTimeMissedCalls = true;
		private bool calcRepeatedMissedAsOne = true;
		private int callbackFirstAttemptMax = 5;
		private int callbackSecondAttemptMax = 20;
		private int callbackThirdAttemptMax = 35;
		private float missedCallsGoodMax = 4.0f;
		private float missedCallsBadMin = 6.0f;
		private float regulationGoodMax = 5.0f;
		private float regulationBadMin = 15.0f;
		
		public FormSettings() {
			InitializeComponent();
			UpdateControls();
		}

		private void UpdateControls() {
			checkBoxCalcRepeatedAsOne.Checked = Properties.Settings.Default.CalcRepeatedMissedAsOne;
			checkBoxInternalNumbers.Checked = Properties.Settings.Default.IgnoreInternalMissedCalls;

			checkBoxWorkingTime.Checked = Properties.Settings.Default.IgnoreNonworkingTimeMissedCalls;
			groupBoxWorkingTime.Enabled = checkBoxWorkingTime.Checked;
			dateTimePickerBegin.Value = new DateTime(2017, 1, 1) + Properties.Settings.Default.WorkingTimeBegin;
			dateTimePickerEnd.Value = new DateTime(2017, 1, 1) + Properties.Settings.Default.WorkingTimeEnd;

			textBoxFirstAttemptMax.Text = Properties.Settings.Default.CallbackFirstAttemptMax.ToString();
			textBoxSecondAttemptMax.Text = Properties.Settings.Default.CallbackSecondAttemptMax.ToString();
			textBoxThirdAttemptMax.Text = Properties.Settings.Default.CallbackThirdAttemptMax.ToString();
			textBoxSecondAttemptMin.Text = textBoxFirstAttemptMax.Text;
			textBoxThirdAttemptMin.Text = textBoxSecondAttemptMax.Text;

			maskedTextBoxMissedGoodMax.Text = GetPercentStringValue(Properties.Settings.Default.MissedCallsGoodMax);
			maskedTextBoxMissedBadMin.Text = GetPercentStringValue(Properties.Settings.Default.MissedCallsBadMin);
			maskedTextBoxMissedMediumMin.Text = maskedTextBoxMissedGoodMax.Text;
			maskedTextBoxMissedMediumMax.Text = maskedTextBoxMissedBadMin.Text;

			maskedTextBoxRegulationGoodMax.Text = GetPercentStringValue(Properties.Settings.Default.RegulationGoodMax);
			maskedTextBoxRegulationBadMin.Text = GetPercentStringValue(Properties.Settings.Default.RegulationBadMin);
			maskedTextBoxRegulationMediumMin.Text = maskedTextBoxRegulationGoodMax.Text;
			maskedTextBoxRegulationMediumMax.Text = maskedTextBoxRegulationBadMin.Text;
		}

		private void buttonReset_Click(object sender, EventArgs e) {
			UpdateSettings(true);
			UpdateControls();
			buttonReset.Enabled = false;
			buttonSave.Enabled = false;
		}

		private void buttonSave_Click(object sender, EventArgs e) {
			UpdateSettings(false);
			buttonSave.Enabled = false;
		}

		private void UpdateSettings(bool toDefault) {
			Properties.Settings.Default.CalcRepeatedMissedAsOne =  toDefault ? 
				calcRepeatedMissedAsOne : checkBoxCalcRepeatedAsOne.Checked;

			Properties.Settings.Default.IgnoreInternalMissedCalls = toDefault ?
				ignoreInternalMissedCalls : checkBoxInternalNumbers.Checked;

			Properties.Settings.Default.IgnoreNonworkingTimeMissedCalls = toDefault ?
				ignoreNonworkingTimeMissedCalls : checkBoxWorkingTime.Checked;

			Properties.Settings.Default.WorkingTimeBegin = toDefault ?
				workingTimeBegin : dateTimePickerBegin.Value.TimeOfDay;

			Properties.Settings.Default.WorkingTimeEnd = toDefault ?
				workingTimeEnd : dateTimePickerEnd.Value.TimeOfDay;

			Properties.Settings.Default.CallbackFirstAttemptMax = toDefault ?
				callbackFirstAttemptMax : GetIntValue(textBoxFirstAttemptMax.Text);

			Properties.Settings.Default.CallbackSecondAttemptMax = toDefault ?
				callbackSecondAttemptMax : GetIntValue(textBoxSecondAttemptMax.Text);

			Properties.Settings.Default.CallbackThirdAttemptMax = toDefault ?
				callbackThirdAttemptMax : GetIntValue(textBoxThirdAttemptMax.Text);

			Properties.Settings.Default.MissedCallsGoodMax = toDefault ?
				missedCallsGoodMax : GetPercentFloatValue(maskedTextBoxMissedGoodMax.Text);

			Properties.Settings.Default.MissedCallsBadMin = toDefault ?
				missedCallsBadMin : GetPercentFloatValue(maskedTextBoxMissedBadMin.Text);

			Properties.Settings.Default.RegulationGoodMax = toDefault ?
				regulationGoodMax : GetPercentFloatValue(maskedTextBoxRegulationGoodMax.Text);

			Properties.Settings.Default.RegulationBadMin = toDefault ?
				regulationBadMin : GetPercentFloatValue(maskedTextBoxRegulationBadMin.Text);

			Properties.Settings.Default.Save();
		}

		private void ControlState_Changed(object sender, EventArgs e) {
			if (sender == checkBoxWorkingTime)
				groupBoxWorkingTime.Enabled = checkBoxWorkingTime.Checked;

			if (sender == textBoxFirstAttemptMax)
				textBoxSecondAttemptMin.Text = textBoxFirstAttemptMax.Text;

			if (sender == textBoxSecondAttemptMax)
				textBoxThirdAttemptMin.Text = textBoxSecondAttemptMax.Text;

			if (sender == maskedTextBoxMissedGoodMax)
				maskedTextBoxMissedMediumMin.Text = maskedTextBoxMissedGoodMax.Text;

			if (sender == maskedTextBoxRegulationBadMin)
				maskedTextBoxRegulationMediumMax.Text = maskedTextBoxRegulationBadMin.Text;

			if (sender == maskedTextBoxMissedBadMin)
				maskedTextBoxMissedMediumMax.Text = maskedTextBoxMissedBadMin.Text;

			if (sender == maskedTextBoxRegulationGoodMax)
				maskedTextBoxRegulationMediumMin.Text = maskedTextBoxRegulationGoodMax.Text;

			buttonSave.Enabled = IsSettingsChanged(false);
			buttonReset.Enabled = IsSettingsChanged(true);
		}

		private string GetPercentStringValue(float number) {
			string retValue = number.ToString();
			if (number < 10.0f)
				retValue = "0" + retValue;
			if (!retValue.Contains("."))
				retValue += ".0";

			return retValue;
		}

		private float GetPercentFloatValue(string number) {
			if (number.Contains("%"))
				number = number.Replace("%", "");

			float retValue = 0.0f;
			float.TryParse(number, out retValue);

			return retValue;
		}

		private bool IsSettingsChanged(bool checkToDefault) {
			if (checkBoxCalcRepeatedAsOne.Checked != (checkToDefault ? 
				calcRepeatedMissedAsOne : Properties.Settings.Default.CalcRepeatedMissedAsOne))
				return true;

			if (checkBoxInternalNumbers.Checked != (checkToDefault ?
				ignoreInternalMissedCalls : Properties.Settings.Default.IgnoreInternalMissedCalls))
				return true;

			if (checkBoxWorkingTime.Checked != (checkToDefault ?
				ignoreNonworkingTimeMissedCalls : Properties.Settings.Default.IgnoreNonworkingTimeMissedCalls))
				return true;

			if (!dateTimePickerBegin.Value.TimeOfDay.Equals(checkToDefault ?
				workingTimeBegin : Properties.Settings.Default.WorkingTimeBegin))
				return true;

			if (!dateTimePickerEnd.Value.TimeOfDay.Equals(checkToDefault ?
				workingTimeEnd : Properties.Settings.Default.WorkingTimeEnd))
				return true;

			if (GetIntValue(textBoxFirstAttemptMax.Text) != (checkToDefault ?
				callbackFirstAttemptMax : Properties.Settings.Default.CallbackFirstAttemptMax))
				return true;

			if (GetIntValue(textBoxSecondAttemptMax.Text) != (checkToDefault ?
				callbackSecondAttemptMax : Properties.Settings.Default.CallbackSecondAttemptMax))
				return true;

			if (GetIntValue(textBoxThirdAttemptMax.Text) != (checkToDefault ?
				callbackThirdAttemptMax : Properties.Settings.Default.CallbackThirdAttemptMax))
				return true;

			if (GetPercentFloatValue(maskedTextBoxMissedGoodMax.Text) != (checkToDefault ?
				missedCallsGoodMax : Properties.Settings.Default.MissedCallsGoodMax))
				return true;

			if (GetPercentFloatValue(maskedTextBoxMissedBadMin.Text) != (checkToDefault ?
				missedCallsBadMin : Properties.Settings.Default.MissedCallsBadMin))
				return true;

			if (GetPercentFloatValue(maskedTextBoxRegulationGoodMax.Text) != (checkToDefault ?
				regulationGoodMax : Properties.Settings.Default.RegulationGoodMax))
				return true;

			if (GetPercentFloatValue(maskedTextBoxRegulationBadMin.Text) != (checkToDefault ?
				regulationBadMin : Properties.Settings.Default.RegulationBadMin))
				return true;

			return false;
		}

		private int GetIntValue(string value) {
			int retValue = 0;
			int.TryParse(value, out retValue);

			return retValue;
		}

		private void FormSettings_FormClosing(object sender, FormClosingEventArgs e) {
			bool haveChanges = IsSettingsChanged(false);
			if (!haveChanges)
				return;

			if (MessageBox.Show(this, "Имеются измененные настройки, Вы уверены, что хотите закрыть окно без сохранения?",
				"Сохранение", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
				return;

			e.Cancel = true;
		}

		private void textBox_KeyPress(object sender, KeyPressEventArgs e) {
			if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
				e.Handled = true;
		}
	}
}
