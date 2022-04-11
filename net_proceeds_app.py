import streamlit as st
import pandas as pd
from datetime import date, datetime
from PIL import Image
from to_excel import inputs_to_excel

# https://share.streamlit.io/streamlit/release-demos/0.84/0.84/streamlit_app.py?page=headliner
# https://docs.streamlit.io/library/advanced-features/session-state#initialization

def main():
    logo = Image.open('freedom_logo.png')

    st.set_page_config(
        page_title='Freedom PM & Sales CMA Form',
        page_icon=logo,
        layout='wide'
    )

    col1, col2, col3 = st.columns([1, 2, 1])
    with col1:
        st.write('')
    with col2:
        st.image(logo)
    with col3:
        st.write('')

    disclaimer_container = st.container()
    password_container = st.container()
    description_container = st.container()
    instruction_container = st.container()
    form_container = st.container()

    with disclaimer_container:
        st.markdown('#### **Disclaimer**')
        st.markdown(
            '''
        *These estimates are not guaranteed and may not include escrows. Escrow balances are reimbursed by the existing lender. Taxes, rents & association dues are pro-rated at settlement. Under Virginia Law, the seller's proceeds may not be available for up to 2 business days following recording of the deed. Seller acknowledges receipt of this statement.*
        '''
        )

    with password_container:
        password_guess = st.text_input('Enter a password to gain access to this app', key='password_guess')
        if password_guess != st.secrets['password']:
            st.stop()

    with description_container:
        with st.expander('App Description'):
            st.markdown(
                '''
                ##### Comparative Market Analysis (CMA) Data Inputs
                This application is used to capture pertinent data points related to a CMA, which will then:
                - Calculate the Estimated Total Net Proceeds
                - Produce the Estimated Total Net Proceeds form in Excel
                ''')

    with instruction_container:
        with st.expander('Instructions'):
            st.markdown(
                '''
                - Enter known data into applicable fields
                - For sliders, move the slider to the general vicinity of the desired value and then use left and right arrows to fine-tune the value
                    - Some of the sliders have values already preset
                    - If you want a different value, just move the slider        
                - When all data is entered, press the 'Calculate Total Net Proceeds' button
                - After pressing the 'Calculate Total Net Proceeds' button: 
                    - Total Net Proceeds value will appear
                    - A new button will appear, 'Download Net Proceeds Form'
                    - Pressing 'Download Net Proceeds Form' button will download an Excel file
                    - The file will appear in the browser downloads location
                - The print area of the Excel file has already been set to allow for easy printing to PDF or paper
                '''
            )

    if 'preparer' not in st.session_state:
        # st.session_state['cma_form'] = False
        # st.session_state['password_guess'] = ''
        st.session_state['preparer'] = ''
        st.session_state['prep_date'] = date.today()
        st.session_state['seller_name'] = ''
        st.session_state['seller_address'] = ''
        st.session_state['estimated_payoff_first_trust'] = 0
        st.session_state['estimated_payoff_second_trust'] = 0
        st.session_state['annual_tax_amt'] = 0
        st.session_state['prorated_tax_amt'] = 0
        st.session_state['annual_hoa_condo_amt'] = 0
        st.session_state['prorated_hoa_condo_amt'] = 0
        st.session_state['rec_list_price'] = 0
        # st.session_state['down_payment_pct'] = 0.0
        st.session_state['closing_subsidy_radio'] = 'Percent of Recommended List Price'
        st.session_state['closing_subsidy_flat_amt'] = 0
        st.session_state['closing_subsidy_pct'] = 0.0
        st.session_state['closing_subsidy_amt'] = 0
        st.session_state['listing_company_pct'] = 0.025
        st.session_state['selling_company_pct'] = 0.025
        st.session_state['processing_fee'] = 0
        st.session_state['settlement_fee'] = 450
        st.session_state['deed_prep_fee'] = 150
        st.session_state['lien_release_fee'] = 100
        st.session_state['lien_trust_qty'] = 1
        st.session_state['recording_release_fee'] = 38
        st.session_state['release_qty'] = 1
        st.session_state['grantors_tax_pct'] = 0.001
        st.session_state['congestion_tax_pct'] = 0.002
        st.session_state['pest_inspection_fee'] = 50
        st.session_state['poa_condo_disclosure_fee'] = 350
        st.session_state['other_fee_name'] = ''
        st.session_state['other_fee_amt'] = 0

    def update_cma_form():
        st.session_state.prep_date = st.session_state.prep_date
        st.session_state.prorated_tax_amt = st.session_state.update_annual_tax_amt / 12 * 3
        st.session_state.prorated_hoa_condo_amt = st.session_state.update_annual_hoa_condo_amt / 12 * 3
        # st.session_state.down_payment_pct = st.session_state.update_down_payment_pct / 100
        st.session_state.closing_subsidy_pct = st.session_state.update_closing_subsidy_pct / 100
        st.session_state.listing_company_pct = st.session_state.update_listing_company_pct / 100
        st.session_state.selling_company_pct = st.session_state.update_selling_company_pct / 100
        st.session_state.grantors_tax_pct = st.session_state.update_grantors_tax_pct / 100
        st.session_state.congestion_tax_pct = st.session_state.update_congestion_tax_pct / 100

        if st.session_state.closing_subsidy_radio == 'Percent of Recommended List Price':
            st.session_state.closing_subsidy_amt = st.session_state.closing_subsidy_pct * st.session_state.rec_list_price
        elif st.session_state.closing_subsidy_radio == 'Flat $ Amount':
            st.session_state.closing_subsidy_amt = st.session_state.closing_subsidy_flat_amt

        st.session_state.housing_costs_subtotal = (
                st.session_state.estimated_payoff_first_trust +
                st.session_state.estimated_payoff_second_trust +
                st.session_state.closing_subsidy_amt +
                st.session_state.prorated_tax_amt +
                st.session_state.prorated_hoa_condo_amt
                )
        st.session_state.brokerage_cost_subtotal = (
                st.session_state.listing_company_pct * st.session_state.rec_list_price +
                st.session_state.selling_company_pct * st.session_state.rec_list_price +
                st.session_state.processing_fee
        )
        st.session_state.closing_cost_subtotal = (
                st.session_state.settlement_fee +
                st.session_state.deed_prep_fee +
                st.session_state.lien_release_fee * st.session_state.lien_trust_qty
        )
        st.session_state.misc_cost_subtotal = (
                st.session_state.recording_release_fee * st.session_state.release_qty +
                st.session_state.grantors_tax_pct * st.session_state.rec_list_price +
                st.session_state.congestion_tax_pct * st.session_state.rec_list_price +
                st.session_state.pest_inspection_fee +
                st.session_state.poa_condo_disclosure_fee +
                st.session_state.other_fee_amt
        )
        st.session_state.total_cost_of_settlement = (
                st.session_state.housing_costs_subtotal +
                st.session_state.brokerage_cost_subtotal +
                st.session_state.closing_cost_subtotal +
                st.session_state.misc_cost_subtotal
        )

        st.session_state.estimated_total_net_proceeds = st.session_state.rec_list_price - st.session_state.total_cost_of_settlement
        # return estimated_total_net_proceeds

    with form_container:
        with st.expander('Data Entry'):
            with st.form(key='cma_form'):
                st.markdown('##### **From Preparation Data**')
                st.text_input('Enter the preparing agent\'s name', key='preparer')
                st.date_input('Enter preparation date of the form', key='prep_date')
                st.write('')
                st.write('---')
                st.write('')

                seller_data, buyer_data = st.columns(2)
                with seller_data:
                    st.markdown('##### **Seller-Specific Data**')
                    st.text_input('Enter seller\'s name(s)', key='seller_name')
                    st.text_input("Enter seller's address", key='seller_address')
                    st.slider("Estimated Payoff - First Trust ($)", 0, 1000000, step=1000, key='estimated_payoff_first_trust')
                    st.slider("Estimated Payoff - Second Trust ($)", 0, 1000000, step=1000, key='estimated_payoff_second_trust')
                    st.slider("Annual Tax Amount ($)", 0, 25000, step=1, key='update_annual_tax_amt')
                    st.slider('Annual HOA / Condo Amount ($)', 0, 10000, step=1, key='update_annual_hoa_condo_amt')

                with buyer_data:
                    st.markdown('##### **Buyer-Specific Data**')
                    st.slider("Recommended Listing Price ($)", 0, 1500000, step=1000, key='rec_list_price')
                    # st.slider('Percent Down Payment (%)', 0.0, 100.0, step=0.01, key='update_down_payment_pct')
                    st.write('')
                    st.write('')
                    st.markdown('*App Default for Closing Cost Subsidy is 0% of Rec. List Price*')
                    st.markdown('*If no Closing Cost Subsidy is requested, leave as-is*')
                    st.markdown('*If a Subsidy is requested, choose appropriate option and adjust associated slider*')
                    st.radio('Closing Cost Subsidy Choice', ['Percent of Recommended List Price', 'Flat $ Amount'], key='closing_subsidy_radio')
                    st.slider('Buyer requests closing cost subsidy of ($):', 0, 100000, step=50, key='closing_subsidy_flat_amt')
                    st.slider('Buyer requests closing cost subsidy of (%):', 0.0, 100.0, step=0.01, key='update_closing_subsidy_pct')
                st.write('')
                st.write('---')
                st.write('')

                brokerage_data, closing_cost_data, misc_data = st.columns(3)
                with brokerage_data:
                    st.markdown('##### **Brokerage Cost Data**')
                    st.slider("Listing Company's Compensation (%)", 0.0, 6.0, 2.5, step=0.01, format='%.2f', key='update_listing_company_pct')
                    st.write('')
                    st.slider("Selling Company's Compensation (%)", 0.0, 6.0, 2.5, step=0.01, format='%.2f', key='update_selling_company_pct')
                    processing_fee = st.slider('Processing Fee Amount ($)', 0, 20000, step=1, key='processing_fee')

                with closing_cost_data:
                    st.markdown('##### **Closing Cost Data**')
                    st.slider('Settlement Fee Amount ($)', 0, 1000, step=1, key='settlement_fee')
                    st.slider('Deed Preparation Fee Amount ($)', 0, 1000, step=1, key='deed_prep_fee')
                    st.slider('Release of Liens / Trusts Fee Amount ($)', 0, 1000, step=1, key='lien_release_fee')
                    st.slider('Number of Liens / Trusts', 0, 10, step=1, key='lien_trust_qty')

                with misc_data:
                    st.markdown('##### **Miscellaneous Cost Data**')
                    st.slider('Recording Release(s) Fee Amount ($)', 0, 250, step=1, key='recording_release_fee')
                    st.slider('Number of Releases', 0, 10, step=1, key='release_qty')
                    st.slider("Grantor's Tax (%)", 0.0, 1.0, 0.1, step=0.01, format='%.2f', key='update_grantors_tax_pct')
                    st.slider("Congestion Relief Tax (%)", 0.0, 1.0, 0.2, step=0.01, format='%.2f', key='update_congestion_tax_pct')
                    st.slider("Pest Inspection Fee Amount ($)", 0, 100, step=1, key='pest_inspection_fee')
                    st.slider("POA / Condo Disclosure Fee Amount ($)", 0, 500, step=1, key='poa_condo_disclosure_fee')
                    st.text_input('Enter name of another fee, if applicable', key='other_fee_name')
                    st.slider('Enter the amount for the \'Other\' fee, if applicable', 0, 100000, step=1000, key='other_fee_amt')

                submit = st.form_submit_button(label='Calculate Total Net Estimated Proceeds', on_click=update_cma_form)

    if submit:
        st.write(f'Calculate Estimated Total Net Proceeds: ${st.session_state.estimated_total_net_proceeds}')

    # st.write(st.session_state)

        proceeds_form = inputs_to_excel(agent=st.session_state.preparer,
                                        date=st.session_state.prep_date,
                                        seller=st.session_state.seller_name,
                                        address=st.session_state.seller_address,
                                        first_trust=st.session_state.estimated_payoff_first_trust,
                                        second_trust=st.session_state.estimated_payoff_second_trust,
                                        annual_taxes=st.session_state.update_annual_tax_amt,
                                        prorated_taxes=st.session_state.prorated_tax_amt,
                                        annual_hoa_condo_amt=st.session_state.annual_hoa_condo_amt,
                                        prorated_annual_hoa_condo_amt=st.session_state.prorated_hoa_condo_amt,
                                        list_price=st.session_state.rec_list_price,
                                        down_payment_pct=st.session_state.down_payment_pct,
                                        closing_subsidy_amt=st.session_state.closing_subsidy_amt,
                                        listing_company_pct=st.session_state.listing_company_pct,
                                        selling_company_pct=st.session_state.selling_company_pct,
                                        processing_fee=st.session_state.processing_fee,
                                        settlement_fee=st.session_state.settlement_fee,
                                        deed_preparation_fee=st.session_state.deed_prep_fee,
                                        lien_release_fee=st.session_state.lien_release_fee,
                                        lien_trust_qty=st.session_state.lien_trust_qty,
                                        recording_release_fee=st.session_state.recording_release_fee,
                                        recording_release_qty=st.session_state.release_qty,
                                        grantors_tax_pct=st.session_state.grantors_tax_pct,
                                        congestion_tax_pct=st.session_state.congestion_tax_pct,
                                        pest_inspection_fee=st.session_state.pest_inspection_fee,
                                        poa_condo_disclosure_fee=st.session_state.poa_condo_disclosure_fee,
                                        other_fee_name=st.session_state.other_fee_name,
                                        other_fee_amt=st.session_state.other_fee_amt)

        st.download_button(
            label='Download Net Proceeds Form',
            data=proceeds_form,
            mime='xlsx',
            file_name=f"net_proceeds_form_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )







if __name__ == '__main__':
    main()