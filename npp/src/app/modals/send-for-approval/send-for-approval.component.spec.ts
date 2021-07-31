import { ComponentFixture, TestBed } from '@angular/core/testing';

import { SendForApprovalComponent } from './send-for-approval.component';

describe('SendForApprovalComponent', () => {
  let component: SendForApprovalComponent;
  let fixture: ComponentFixture<SendForApprovalComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ SendForApprovalComponent ]
    })
    .compileComponents();
  });

  beforeEach(() => {
    fixture = TestBed.createComponent(SendForApprovalComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
