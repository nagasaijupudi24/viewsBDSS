import * as React from "react";
import styles from "../styles/passcodePage.module.scss";
import * as CryptoJS from "crypto-js";
import { IXenWpUcoBankProps } from "../IXenWpUcoBankProps";
import {
  PrimaryButton,
} from "@fluentui/react";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "bootstrap/dist/css/bootstrap.min.css";
import "bootstrap/dist/js/bootstrap.bundle.min.js";
import "@pnp/sp/site-users/web";
import "../CustomStyles/Custom.css";
import {
  IHttpClientOptions,
  HttpClient,
} from "@microsoft/sp-http";
import { Dialog, DialogType, DialogFooter } from "@fluentui/react/lib/Dialog";
export interface IPasscodeState {
  passcode: string;
  confimPasscode: string;
  otp: any;
  passcodeSent: boolean;
  passcodeVerified: boolean;
  error: string;
  timer: number;
  timerExpired: boolean;
  timerId: any;
  verifyContentHide: boolean;
  saveButtonVisible: boolean;
  showSuccessPopup: boolean;
  isPasscodeVisible: boolean;
  isConfirmPasscodeVisible: boolean;
  showErrorPopup: boolean;
  encryptedPasscode: any;
  otpSent: any;
}

const dragOptions = {
  moveMenuItemText: "Move",
  closeMenuItemText: "Close",
};
const modalPropsStyles = { main: {
  maxWidth:600
 } };
const dialogContentProps = {
  type: DialogType.normal,
  title: "Alert",
};
export default class PasscodePage extends React.Component<
  IXenWpUcoBankProps,
  IPasscodeState,
  {}
> {
  private _inputsRef: any;

  private timer: any;
  private _listName: any;
  constructor(props: IXenWpUcoBankProps) {
    super(props);
    this.state = {
      passcode: "",
      confimPasscode: "",
      otp: new Array(6).fill(""),
      passcodeSent: false,
      passcodeVerified: false,
      error: "",
      timer: 180,
      timerExpired: false,
      timerId: null,
      verifyContentHide: true,
      saveButtonVisible: false,
      showSuccessPopup: true,
      isPasscodeVisible: false,
      isConfirmPasscodeVisible: false,
      showErrorPopup: true,
      encryptedPasscode: "",
      otpSent: null,
    };
    const listObj = this.props.listName;
    this._listName = listObj?.title;
    this._inputsRef = [];
  }
 /*  // Function to generate a random 6-digit OTP */
  private _generateOTP = () => {
    const array = new Uint32Array(1);
  
    window.crypto.getRandomValues(array);
    const otp = array[0] % 900000 + 100000; // Ensures a 6-digit OTP
    return otp;
};


 /*  // Function to handle passcode submission */
  private _handleAddPasscode = async (event: any) => {
    event.preventDefault();
    const { passcode, confimPasscode } = this.state;
    const regex = /^(?=.*[A-Za-z])(?=.*\d)[A-Za-z\d]+$/;
    if (passcode.length < 6 || confimPasscode.length < 6) {
      this.setState({
        error: "Passcode should contain 6-characters.",
      });
    } else if (!passcode || !confimPasscode) {
      this.setState({
        error: "Passcode and confirm passcode should not be empty.",
      });

      return;
    } else if (!regex.test(passcode) || !regex.test(confimPasscode)) {
      this.setState({
        error: "Passcode should contain both Alphabets and Numbers.",
      });
      return;
    } else if (passcode !== confimPasscode) {
      this.setState({
        error: "Passcode and confirm passcode should match.",
      });
      return;
    } else {
      const otp = this._generateOTP();
      this._createOtp(otp);
    }
  };

 
  // Function to verify/*  */ the passcode entered by the user
  private _handleVerifyPasscode = async () => {
    clearInterval(this.state.timerId); /* // Stop the timer */
    // setVerifyContentHide(false); // Hide the/*  */ timer text when verify button is clicked
    const enteredOtp = this.state.otp.join("");
    if (enteredOtp.length !== 6) {
      this.setState({
        error: "Please enter the 6-digit verification code.",
        showErrorPopup: false,
      });
    }
   /*  // Compare decrypted passcode with user entered OTP */
    else if (enteredOtp === this.state.otpSent.toString()) {
      this.setState({
        passcodeVerified: true,
        saveButtonVisible: true,
        verifyContentHide: false,
      });
    } else {
      this.setState({
        showErrorPopup: false,
        error: "Incorrect passcode entered please try again.",
      });
    }
  };

 /*  // Function to save the passcode to local storage */
  private _handleSavePasscode = async () => {
  /*   // Define key and IV */
    const key = CryptoJS.enc.Utf8.parse("b75524255a7f54d2726a951bb39204df");
    const iv = CryptoJS.enc.Utf8.parse("1583288699248111");
    const text = this.state.passcode; /* // Your passcode input */

 /*    // Encrypt the passcode */
    const encryptedCP = CryptoJS.AES.encrypt(text, key, { iv: iv });
    const cryptText = encryptedCP.toString();

/*     Decrypt the encrypted text directly
    const decryptedWA = CryptoJS.AES.decrypt(encryptedCP, key, { iv: iv });
    console.log("Decrypted Text (from encryptedCP): ", decryptedWA.toString(CryptoJS.enc.Utf8));

    Alternatively: Decrypting from the base64 string
    const decryptedFromText = CryptoJS.AES.decrypt(cryptText, key, { iv: iv });
    console.log("Decrypted Text (from cryptText): ", decryptedFromText.toString(CryptoJS.enc.Utf8)); */
    try {
      let user = await this.props.sp?.web.currentUser();
      const items = await this.props.sp.web.lists
        .getByTitle(this._listName)
        .items.filter(`AuthorId eq ${user.Id}`)();
      if (items && items.length > 0) {
        await this.props.sp.web.lists
          .getByTitle(this._listName)
          .items.getById(items[0].Id)
          .update({
            passcode: cryptText,
          });
      } else {
        await this.props.sp.web.lists.getByTitle(this._listName).items.add({
          passcode: cryptText,
          UserId: user.Id,
          Title: user.Title,
        });
      }
      this.setState({ showSuccessPopup: !this.state.showSuccessPopup });

    } catch (err) {
      console.error(err);
    }
  };

 /*  // Function to handle closing the success popup */
  private _handleCloseSuccessPopup = () => {
    this.setState({
      showSuccessPopup: true,
      saveButtonVisible: false,
      passcode: "",
      confimPasscode: "",
      otp: new Array(6).fill(""),
      passcodeSent: false,
      passcodeVerified: false,
      error: "",
      timer: 180,
      timerExpired: false,
      verifyContentHide: true,
    });
  };

 /*  // Function to resend the passcode */
  private _handleResendPasscode = async (event: any) => {
    clearInterval(this.timer); /* // Clear any existing interval */
/*     // Reset timer to 3 minutes */
    this.setState({
      otp: new Array(6).fill(""),
      timerExpired: false,
      timer: 180,
      passcodeSent: true,
      passcodeVerified: false,
    });
    this._handleAddPasscode(event);
  };

/*   // Handle input change in OTP fields */
  private _handleOtpChange = (index: number, value: string) => {
    if (/[^a-zA-Z0-9]/.test(value)) return;
/*     // Only allow alphanumeric input */
    const newOtp = [...this.state.otp];
    newOtp[index] = value;
    this.setState({
      otp: newOtp,
    });
    console.log(this._inputsRef, "this._inputsRef");

/*     // Focus the next input field */
    if (value !== "" && index < 5) {
      this._inputsRef[index + 1].focus();
    }
  };

/*   // Handle keydown event for backspace and arrow keys */
  private _handleKeyDown = (index: number, event: { key: string }) => {
    if (
      event.key === "Backspace" &&
      this.state.otp[index] === "" &&
      index > 0
    ) {
      this._inputsRef[index - 1].focus();
    }
    if (event.key === "ArrowLeft" && index > 0) {
      this._inputsRef[index - 1].focus();
    }
    if (event.key === "ArrowRight" && index < 5) {
      this._inputsRef[index + 1].focus();
    }
  };

  
  private _renderOtpInputs = () => {/* // Render OTP input boxes */
    return this.state.otp.map((data: any, index: any) => (
      <input
        key={index+1}
        type="text"
        maxLength={1}
        value={this.state.otp[index]}
        onChange={(e) => this._handleOtpChange(index, e.target.value)}
        onKeyDown={(e) => this._handleKeyDown(index, e)}
        ref={(el) => (this._inputsRef[index] = el)}
        className={styles.otp_inputfield}
      />
    ));
  };


  private _handlePasscodeChange = (event: { target: { value: any } }) => {
    const passcodeValue = event.target.value;
    const { confimPasscode } = this.state;
    const alphanumericRegex = /^[a-zA-Z0-9]*$/;
    const regex = /^(?=.*[A-Za-z])(?=.*\d)[A-Za-z\d]+$/;
    if (alphanumericRegex.test(passcodeValue)) {
      this.setState({
        passcode: passcodeValue,
      });
    }

    if (confimPasscode !== passcodeValue && confimPasscode !== "") {/* Check if the confirm passcode matches the new passcode value */
      this.setState({
        error: "Passcodes do not match",
      });
    } else if (!regex.test(passcodeValue) && passcodeValue.length > 5) {
      this.setState({
        error: "Passcode should contain both Alphabets and Numbers.",
      });
    } else {
      this.setState({
        error: "",
      });
    }
  };


  private _handleConfirmPasscodeChange = (event: {  /* // Function to handle confirm passcode input change */
    target: { value: any };
  }) => {
    const confirmPasscodeValue = event.target.value;
    const { passcode } = this.state;
    const alphanumericRegex = /^[a-zA-Z0-9]*$/;

    if (alphanumericRegex.test(confirmPasscodeValue)) {
      this.setState({
        confimPasscode: confirmPasscodeValue,
      });
    }
  
    const regex = /^(?=.*[A-Za-z])(?=.*\d)[A-Za-z\d]+$/; /*  // Check if the confirm passcode matches the passcode value
 */
    if (passcode !== confirmPasscodeValue && passcode !== "") {
      this.setState({
        error: "Entered Passcodes do not match",
      });
    } else if (
      !regex.test(confirmPasscodeValue) &&
      confirmPasscodeValue.length > 5
    ) {
      this.setState({
        error: "Passcode should contain both Alphabets and Numbers.",
      });
    } else {
      this.setState({
        error: "",
      });
    }
  };

  private _startTimer = (): void => {
    if (!this.state.timerExpired) {
      this.timer = setInterval(() => { /*  // Start the timer */
        this.setState(
          (prevState) => ({
            timer: prevState.timer - 1,
          }),
          () => {
           
            if (this.state.timer === 0) {/*  // Handle the case when time expires */
              clearInterval(this.timer);
              this.setState({ timerExpired: true });
            }
          }
        );
      }, 1000);
    }
  };
  componentWillUnmount(): void {
    if (this.timer) {
      clearInterval(this.timer);/*  // Clear the timer when the componentunmounts to avoid memory leaks */ 
    }
  }
  private _createOtp = async (otp: any) => {
    const currentUser = await this.props.sp.web.currentUser();
    const params = {
      mailid: currentUser.Email,
      OTP: otp.toString(),
    };
    const postURL = this.props.httpUrl;
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const httpClientOptions: IHttpClientOptions = {
      body: JSON.stringify(params),
      headers: requestHeaders,
    };
    return this.props.context.httpClient
      .post(postURL, HttpClient.configurations.v1, httpClientOptions)
      .then((response: any): Promise<any> => {

        if (response.ok) {
          this.setState({
            passcodeSent: true,
            otpSent: otp,
          });
          this._startTimer();
        }

        return response.status;
      });
  };

  public render(): React.ReactElement<IXenWpUcoBankProps> {
    const {
      passcodeSent,
      passcodeVerified,
      isPasscodeVisible,
      passcode,
      isConfirmPasscodeVisible,
      error,
      timerExpired,
      saveButtonVisible,
      verifyContentHide,
      showSuccessPopup,
      confimPasscode,
    } = this.state;
    const modalProps: any = {
      isBlocking: true,
      styles: modalPropsStyles,
      dragOptions: dragOptions,
    };

    return (
      <section className={styles.XenPasscode}>
        <div className={styles.passcode_header_conatiner}>
          Create/Reset Passcode
        </div>
        <div className={` ${styles._createCard} card w-50 mb-3`}>
          <div className="card-body">
            <div className="card-text">
              {passcodeSent === false && passcodeVerified === false && (
                 <>
                  <div className={styles.passcodeHint}>
                    Your Passcode must consist of 6 alpha-numeric characters
                  </div>
                  <div>
                    <label htmlFor={"passcode"}>Enter a Passcode:</label>
                    <input
                      type={isPasscodeVisible ? "text" : "password"}
                      id="passcode"
                      value={passcode}
                      onChange={this._handlePasscodeChange}
                      className={styles.passcode_input}
                      maxLength={6}
                      pattern="\d*"
                      title="Please enter 6-characters, Combination of Alphabets and Numbers"
                    /> 
            
                  </div>
                  <div>
                    <label htmlFor="confirm_passcode">Confirm Passcode:</label>
                    <div className="passcode-input-container">
                      <input
                        type={isConfirmPasscodeVisible ? "text" : "password"}
                        id="confirm_passcode"
                        value={confimPasscode}
                        onChange={this._handleConfirmPasscodeChange}
                        className={styles.passcode_input}
                        maxLength={6}
                        pattern="\d*"
                        title="Please enter 6-characters, Combination of Alphabets and Numbers"
                      />
                    </div>
                    {error && <p className={styles.confirmerror}>{error}</p>}
                  </div>

                  <div className={styles.passcode_btn_container}>
                    <PrimaryButton onClick={this._handleAddPasscode} iconProps={{iconName:"Send"}}>
                      Submit
                    </PrimaryButton>
                  </div>
                </>
              )}
              {passcodeSent === true && passcodeVerified === false && (
                <>
                  <p>
                    OTP has been sent to your mobile number & email address.
                  </p>
                  <p>Please enter your OTP</p>
                  <div>{this._renderOtpInputs()}</div>

                  {!timerExpired && (
                    <>
                      {!saveButtonVisible && (
                        <>
                          <br />
                          <div className={styles.passcode_btn_container}>
                            <PrimaryButton onClick={this._handleVerifyPasscode} iconProps={{iconName:"Completed"}}>
                              Verify OTP
                            </PrimaryButton>
                          </div>
                        </>
                      )}
                      {verifyContentHide && (
                        <p className={styles.timer}>
                          Time remaining: <b>{this.state.timer}</b> seconds
                        </p>
                      )}
                    </>
                  )}
                  {timerExpired && (
                    <div>
                      <p className={styles.resendPara}>OTP has been expired.</p>
                      <div className={styles.passcode_btn_container}>
                        <PrimaryButton onClick={this._handleResendPasscode} iconProps={{iconName:"Send"}}>
                          Resend OTP
                        </PrimaryButton>
                      </div>
                    </div>
                  )}
                </>
              )}
              {passcodeVerified && (
                <>
                  <p className="otpverified">OTP verified successfully!</p>
                  {saveButtonVisible && (
                    <div className={styles.passcode_btn_container}>
                      <PrimaryButton onClick={this._handleSavePasscode} iconProps={{iconName:"Save"}}>
                        Save Passcode
                      </PrimaryButton>
                    </div>
                  )}
                </>
              )}
            </div>
          </div>
        </div>
        <Dialog
          hidden={this.state.showSuccessPopup}
          onDismiss={() =>
            this.setState({ showSuccessPopup: !showSuccessPopup })
          }
          dialogContentProps={dialogContentProps}
          modalProps={modalProps}
          maxWidth={600}
        >
          <div  className="dialogcontent_">
          <p>Passcode saved successfully!</p>
          </div>
          <DialogFooter>
            <PrimaryButton onClick={this._handleCloseSuccessPopup}   iconProps={{iconName:"ReplyMirrored"}}>
              Ok
            </PrimaryButton>
          </DialogFooter>
        </Dialog>
        <Dialog
          hidden={this.state.showErrorPopup}
          onDismiss={() =>
            this.setState({ showErrorPopup: !this.state.showErrorPopup })
          }
          dialogContentProps={dialogContentProps}
          modalProps={modalProps}
          maxWidth={600}
        >
           <div className="dialogcontent_">
 <p >{this.state.error}</p>
          </div> 
         
          <DialogFooter>
            <PrimaryButton
            iconProps={{iconName:"ReplyMirrored"}}
              onClick={() =>
                this.setState({ showErrorPopup: !this.state.showErrorPopup })
              }
            >
              Ok
            </PrimaryButton>
          </DialogFooter>
        </Dialog>
      </section>
    );
  }
}
