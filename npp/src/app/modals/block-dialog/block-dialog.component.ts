import { Component, Inject, OnInit } from '@angular/core';
import { MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';

@Component({
  selector: 'app-block-dialog',
  templateUrl: './block-dialog.component.html',
  styleUrls: ['./block-dialog.component.scss']
})
export class BlockDialogComponent implements OnInit {

  wait: string = "Update in progress, please do not navigate away from this screen or close the application.";
  message: string = '';

  constructor(@Inject(MAT_DIALOG_DATA) public data: any, private dialogRef: MatDialogRef<BlockDialogComponent>) {
    if (data) {
      this.message = data.message || this.message;
      this.wait = data.wait || this.wait;
    }
  }

  ngOnInit(): void {
    this.dialogRef.disableClose = true;
  }

  onConfirm(): void {
    this.dialogRef.close(true);
  }

  onDismiss(): void {
    this.dialogRef.close(false);
  }
}
