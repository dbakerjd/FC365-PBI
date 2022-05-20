import { Component, Inject, OnInit } from '@angular/core';
import { MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';

@Component({
  selector: 'app-confirm-dialog',
  templateUrl: './confirm-dialog.component.html',
  styleUrls: ['./confirm-dialog.component.scss']
})
export class ConfirmDialogComponent implements OnInit {

  message: string = "Are you sure?";
  confirmButtonText = "Yes";
  cancelButtonText = "Cancel";
  reference: any = undefined; // Used if you need to know the reference of what we are confirming

  constructor(@Inject(MAT_DIALOG_DATA) public data: any, private dialogRef: MatDialogRef<ConfirmDialogComponent>) {
    if (data) {
      this.message = data.message || this.message;
      this.confirmButtonText = data.confirmButtonText || this.confirmButtonText;
      this.cancelButtonText = data.cancelButtonText || this.cancelButtonText;
      this.reference = data.reference || this.reference;
    }
  }

  ngOnInit(): void {
  }

  onConfirm(): void {
    if (this.reference) {
      this.dialogRef.close({
        result: true,
        reference: this.reference
      });
    } else {
      this.dialogRef.close(true);
    } 
  }

  onDismiss(): void {
    if (this.reference) {
      this.dialogRef.close({
        result: false,
        reference: this.reference
      });
    } else {
      this.dialogRef.close(false);
    } 
  }
}
